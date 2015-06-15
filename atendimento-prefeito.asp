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
		body {
			padding-top: 60px !important;
		}

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

		#modalQtdObras {
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
			centerMap();
		}

		function centerMap(){
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
							'Ficha Técnica'+
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
						icon: 'img/point_'+ item.classStatus +'.png',
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
							link = "ficha-tecnica-obra-potencial.asp?cod_empreendimento=" + item.cod_empreendimento;
							break;
					}

					var btnInformacoes = "";
					var btnZoom = '<a href="#" class="btn btn-sm btn-block btn-primary btn-zoom-in" data-latitude="'+ latitude +'" data-longitude="'+ longitude +'">Zoom na Obra</a>';
					var btnVoltarZoom = '<a href="#" class="btn btn-sm btn-block btn-primary btn-zoom-out hide">Voltar</a>';

					if(link != "")
						btnInformacoes = '<br/><br/><a href="'+ link +'" class="btn btn-sm btn-block btn-info btn-dados-municipio">Ficha Técnica</a>'

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
			$("table.table-localidades tbody tr").remove();
			$.each(localidades, function(i, item) {
				var link = "";

				switch(item.cod_situacao_externa){
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
						link = "ficha-tecnica-obra-potencial.asp?cod_empreendimento=" + item.cod_empreendimento;
						break;
				}

				var itemLayout = '<tr>'+
									'<td class="text-middle">'+ item.nme_municipio + '</td>'+
									'<td class="text-middle">'+ item.nme_empreendimento + '</td>'+
									'<td class="text-middle"><a href="'+ link +'&canClose=1" class="btn btn-sm btn-primary" target="_blank">Ficha Técnica</a></td>'+
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
					if(!$(".form-group.result").hasClass("hide"))
						$(".form-group.result").addClass("hide");

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
						$("#btn-clear")
						$(".form-group.result").removeClass("hide");
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
								dsc_situacao_externa: item.dsc_situacao_externa
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
									case 39: // concluídas
										markerData.classStatus = "light_green";
										break;
									case 40: // em atendimento
										markerData.classStatus = "dark_green";
										break;
									case 42: // finalizada e não operando
										markerData.classStatus = "orange";
										break;
									case 43: // paralizada
										markerData.classStatus = "purple";
										break;
									case 44: // atendimento programado
										markerData.classStatus = "light_blue";
										break;
									case 41: // não atendidas
										markerData.classStatus = "red";
										break;
									case 45: // atendimento potencial
										markerData.classStatus = "dark_blue";
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

		function hideCounterButtons() {
			$(".fa-spin").removeClass("hide");
			$(".btn-details").addClass("hide");
		}

		function showCounterButtons() {
			$.ajax({
				url: "query-to-json-util.asp",
				method: "POST",
				data: {
					sql: "SELECT * FROM c_conta_situacao"
				},
				beforeSend: function() {
					hideCounterButtons();
				},
				success: function(data, textStatus, jqXHR){
					data = JSON.parse(data);

					if(data.length > 0) {
						$(".btn-show-obras-concluidas").text(data[0].qtd_concluida + " Obra(s)");
						$(".btn-show-obras-em-andamento").text(data[0].qtd_em_andamento + " Obra(s)");
						
						$(".btn-show-obras-em-atendimento").text(data[0].qtd_em_atendimento + " Obra(s)");
						$(".btn-show-obras-finalizada-nao-operando").text(data[0].qtd_finalizada_nao_operando + " Obra(s)");
						$(".btn-show-obras-paralizadas").text(data[0].qtd_paralizadas + " Obra(s)");

						$(".btn-show-obras-atendimento-programado").text(data[0].qtd_atendimento_programado + " Obra(s)");
						$(".btn-show-obras-nao-atendidas").text(data[0].qtd_nao_atendidas + " Obra(s)");
						$(".btn-show-obras-atendimento-potencial").text(data[0].qtd_atendimento_potencial + " Obra(s)");

						addButtonDetailsClickListener();
					}

					$(".fa-spin").addClass("hide");
					$(".btn-details").removeClass("hide");
				},
				error: function(jqXHR, textStatus, errorThrown){
					console.log(jqXHR, textStatus, errorThrown);
				}
			});
		}

		function addButtonDetailsClickListener() {
			$(".btn-details").on("click", function(){
				if($(this).hasClass("btn-show-obras-concluidas")){
					$("#modalQtdObrasLabel.modal-title").text("Obras concluídas");
					loadEmpreendimentosBySituacao(""+
						"SELECT tb_pi.PI, c_lista_predios.nme_municipio, tb_pi.nome_empreendimento, tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao AS dsc_situacao_externa "+ 
						"FROM (tb_pi INNER JOIN c_lista_predios ON tb_pi.id_predio = c_lista_predios.id_predio) "+
						"LEFT JOIN tb_situacao_pi ON tb_pi.cod_situacao_externa = tb_situacao_pi.cod_situacao "+
						"WHERE tb_situacao_pi.cod_situacao = 39 "+
						"ORDER BY nme_municipio ASC, nome_empreendimento ASC");
				}
				else if($(this).hasClass("btn-show-obras-em-andamento")){
					$("#modalQtdObrasLabel.modal-title").text("Obras em Andamento");
					loadEmpreendimentosBySituacao(""+
						"SELECT tb_pi.PI, c_lista_predios.nme_municipio, tb_pi.nome_empreendimento, tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao AS dsc_situacao_externa "+ 
						"FROM (tb_pi INNER JOIN c_lista_predios ON tb_pi.id_predio = c_lista_predios.id_predio) "+
						"LEFT JOIN tb_situacao_pi ON tb_pi.cod_situacao_externa = tb_situacao_pi.cod_situacao "+
						"WHERE tb_situacao_pi.cod_situacao IN (40,42,43) "+
						"ORDER BY nme_municipio ASC, nome_empreendimento ASC");
				}
				else if($(this).hasClass("btn-show-obras-em-atendimento")){
					$("#modalQtdObrasLabel.modal-title").text("Obras em Atendimento");
					loadEmpreendimentosBySituacao(""+
						"SELECT tb_pi.PI, c_lista_predios.nme_municipio, tb_pi.nome_empreendimento, tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao AS dsc_situacao_externa "+ 
						"FROM (tb_pi INNER JOIN c_lista_predios ON tb_pi.id_predio = c_lista_predios.id_predio) "+
						"LEFT JOIN tb_situacao_pi ON tb_pi.cod_situacao_externa = tb_situacao_pi.cod_situacao "+
						"WHERE tb_situacao_pi.cod_situacao = 40 "+
						"ORDER BY nme_municipio ASC, nome_empreendimento ASC");
				}
				else if($(this).hasClass("btn-show-obras-finalizada-nao-operando")){
					$("#modalQtdObrasLabel.modal-title").text("Obras Finalizadas e Não Operando");
					loadEmpreendimentosBySituacao(""+
						"SELECT tb_pi.PI, c_lista_predios.nme_municipio, tb_pi.nome_empreendimento, tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao AS dsc_situacao_externa "+ 
						"FROM (tb_pi INNER JOIN c_lista_predios ON tb_pi.id_predio = c_lista_predios.id_predio) "+
						"LEFT JOIN tb_situacao_pi ON tb_pi.cod_situacao_externa = tb_situacao_pi.cod_situacao "+
						"WHERE tb_situacao_pi.cod_situacao = 42 "+
						"ORDER BY nme_municipio ASC, nome_empreendimento ASC");
				}
				else if($(this).hasClass("btn-show-obras-paralizadas")){
					$("#modalQtdObrasLabel.modal-title").text("Obras Paralizadas");
					loadEmpreendimentosBySituacao(""+
						"SELECT tb_pi.PI, c_lista_predios.nme_municipio, tb_pi.nome_empreendimento, tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao AS dsc_situacao_externa "+ 
						"FROM (tb_pi INNER JOIN c_lista_predios ON tb_pi.id_predio = c_lista_predios.id_predio) "+
						"LEFT JOIN tb_situacao_pi ON tb_pi.cod_situacao_externa = tb_situacao_pi.cod_situacao "+
						"WHERE tb_situacao_pi.cod_situacao = 43 "+
						"ORDER BY nme_municipio ASC, nome_empreendimento ASC");
				}
				else if($(this).hasClass("btn-show-obras-atendimento-programado")){
					$("#modalQtdObrasLabel.modal-title").text("Obras Atendimento Programado");
					loadEmpreendimentosBySituacao(""+
						"SELECT tb_pi.PI, c_lista_predios.nme_municipio, tb_pi.nome_empreendimento, tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao AS dsc_situacao_externa "+ 
						"FROM (tb_pi INNER JOIN c_lista_predios ON tb_pi.id_predio = c_lista_predios.id_predio) "+
						"LEFT JOIN tb_situacao_pi ON tb_pi.cod_situacao_externa = tb_situacao_pi.cod_situacao "+
						"WHERE tb_situacao_pi.cod_situacao = 44 "+
						"ORDER BY nme_municipio ASC, nome_empreendimento ASC");
				}
				else if($(this).hasClass("btn-show-obras-nao-atendidas")){
					$("#modalQtdObrasLabel.modal-title").text("Obras Não Atendidas");
					loadEmpreendimentosBySituacao(""+
						"SELECT tb_pi.PI, c_lista_predios.nme_municipio, tb_pi.nome_empreendimento, tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao AS dsc_situacao_externa "+ 
						"FROM (tb_pi INNER JOIN c_lista_predios ON tb_pi.id_predio = c_lista_predios.id_predio) "+
						"LEFT JOIN tb_situacao_pi ON tb_pi.cod_situacao_externa = tb_situacao_pi.cod_situacao "+
						"WHERE tb_situacao_pi.cod_situacao = 41 "+
						"ORDER BY nme_municipio ASC, nome_empreendimento ASC");
				}
				else if($(this).hasClass("btn-show-obras-atendimento-potencial")){
					$("#modalQtdObrasLabel.modal-title").text("Obras Atendimento Potencial");
					loadEmpreendimentosBySituacao(""+
						"SELECT tb_pi.PI, c_lista_predios.nme_municipio, tb_pi.nome_empreendimento, tb_situacao_pi.cod_situacao, tb_situacao_pi.desc_situacao AS dsc_situacao_externa "+ 
						"FROM (tb_pi INNER JOIN c_lista_predios ON tb_pi.id_predio = c_lista_predios.id_predio) "+
						"LEFT JOIN tb_situacao_pi ON tb_pi.cod_situacao_externa = tb_situacao_pi.cod_situacao "+
						"WHERE tb_situacao_pi.cod_situacao = 45 "+
						"ORDER BY nme_municipio ASC, nome_empreendimento ASC");
				}
			});
		}

		function loadEmpreendimentosBySituacao(sql) {
			$.ajax({
				url: "query-to-json-util.asp",
				method: "POST",
				data: {
					sql: sql
				},
				beforeSend: function() {
					$("#modalLoading").modal("show");
				},
				success: function(data, textStatus, jqXHR){
					data = JSON.parse(data);

					if(data.length > 0) {
						var items = [];
						
						$.each(data, function(i, item) {

							var itemData = {
								cod_empreendimento: item.PI,
								nme_municipio: item.nme_municipio,
								nme_empreendimento: item.nome_empreendimento,
								cod_situacao_externa: item.cod_situacao,
								dsc_situacao_externa: item.dsc_situacao_externa
							};

							items.push(itemData);
						});

						listEmpreendimentosTable(items);

						$("#modalQtdObras").modal("show");
					}

					$("#modalLoading").modal("hide");
				},
				error: function(jqXHR, textStatus, errorThrown){
					console.log(jqXHR, textStatus, errorThrown);
				}
			});
		}

		function createAlfaLinks() {
			for (var i = 0; i != 26; ++i){
				alfa = String.fromCharCode(i + 65);
				tmp = '<a href="#" class="alfa-link-item" data-alfa-letter="'+ alfa +'">'+ alfa +'</a>&nbsp;&nbsp;';
				$(".alfa-links").append(tmp);
			}
			addAlfaItemEventListener();
		}

		function addAlfaItemEventListener() {
			$("a.alfa-link-item").on("click", function(){
				$("#txt-municipio").val($(this).data().alfaLetter);
				$("#form-search").submit();
			});
		}

		function loadAllMarkers() {
			loadEmpreendimentos(sqlFiltroSituacao + " AND cod_situacao = 39 OR cod_situacao IN (40,42,43) OR cod_situacao = 44 OR cod_situacao = 41 OR cod_situacao = 45");

			$.each($("input[type=checkbox]"), function(i, item){
			   $(this).trigger("click");
			});
		}

		var sqlFiltroSituacao = "SELECT tb_pi.PI, tb_pi.nome_empreendimento, tb_pi.municipio, tb_pi.qtd_populacao_urbana_2010, tb_pi.latitude_longitude, tb_pi.cod_situacao FROM tb_pi WHERE 1 = 1 ";
		var hasAndSqlQuery = false;

		$(function() {
			setLayersHeight();
			initMap();

			//createAlfaLinks();

			loadAllMarkers();
			showCounterButtons();

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

			$("#chk-em-atendimento").on("change", function() {
				if((!this.checked) && (!$("#chk-finalizada-nao-operando")[0].checked) && (!$("#chk-paralizadas")[0].checked))
					$("#chk-em-andamento")[0].checked = false;
				else
					$("#chk-em-andamento")[0].checked = true;
			});

			$("#chk-finalizada-nao-operando").on("change", function() {
				if((!this.checked) && (!$("#chk-em-atendimento")[0].checked) && (!$("#chk-paralizadas")[0].checked))
					$("#chk-em-andamento")[0].checked = false;
				else
					$("#chk-em-andamento")[0].checked = true;
			});

			$("#chk-paralizadas").on("change", function() {
				if((!this.checked) && (!$("#chk-em-atendimento")[0].checked) && (!$("#chk-finalizada-nao-operando")[0].checked))
					$("#chk-em-andamento")[0].checked = false;
				else
					$("#chk-em-andamento")[0].checked = true;
			});

			$("#chk-em-andamento").on("change", function() {
				$("#chk-em-atendimento")[0].checked 			= (this.checked);
				$("#chk-finalizada-nao-operando")[0].checked 	= (this.checked);
				$("#chk-paralizadas")[0].checked 				= (this.checked);
			});

			$("#btn-load-situacoes").on("click", function(){
				if(!$(".alert-situacao").hasClass("hide"))
					$(".alert-situacao").addClass("hide");

				var chk_concluidas 				= $("#chk-concluidas");
				var chk_em_andamento 			= $("#chk-em-andamento");

				var chk_em_atendimento 			= $("#chk-em-atendimento");
				var chk_finalizada_nao_operando	= $("#chk-finalizada-nao-operando");
				var chk_paralizadas				= $("#chk-paralizadas");

				var chk_atendimento_programado 	= $("#chk-atendimento-programado");
				var chk_nao_atendida 			= $("#chk-nao-atendida");
				var chk_atendimento_potencial 	= $("#chk-atendimento-potencial");

				var context = "";

				sqlFiltroSituacao = "SELECT tb_pi.PI, tb_pi.nome_empreendimento, tb_pi.municipio, tb_pi.qtd_populacao_urbana_2010, tb_pi.latitude_longitude, tb_pi.cod_situacao FROM tb_pi WHERE 1 = 1 ";
				hasAndSqlQuery = false;

				if($(chk_concluidas)[0].checked) {
					if(!hasAndSqlQuery){
						hasAndSqlQuery = true;
						context = " AND cod_situacao = 39";
					}
					else
						context = " OR cod_situacao = 39";

					sqlFiltroSituacao += context;
				}
				else {
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

				/*context = "";
				if($(chk_em_andamento)[0].checked) {
					if(!hasAndSqlQuery){
						hasAndSqlQuery = true;
						context = " AND cod_situacao IN (40,42,43)";
					}
					else
						context = " OR cod_situacao IN (40,42,43)";

					sqlFiltroSituacao += context;
				}
				else {
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
				}*/

				context = "";
				if($(chk_em_atendimento)[0].checked) {
					if(!hasAndSqlQuery){
						hasAndSqlQuery = true;
						context = " AND cod_situacao = 40";
					}
					else
						context = " OR cod_situacao = 40";

					sqlFiltroSituacao += context;
				}
				else {
					if(sqlFiltroSituacao.indexOf(" AND cod_situacao = 40") != -1)
						context = " AND cod_situacao = 40";
					else if(sqlFiltroSituacao.indexOf(" OR cod_situacao = 40") != -1)
						context = " OR cod_situacao = 40";

					sqlFiltroSituacao = sqlFiltroSituacao.replace(context, "");

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						sqlFiltroSituacao = sqlFiltroSituacao.replace("OR", "AND");
					}

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						hasAndSqlQuery = false;
					}
				}

				context = "";
				if($(chk_finalizada_nao_operando)[0].checked) {
					if(!hasAndSqlQuery){
						hasAndSqlQuery = true;
						context = " AND cod_situacao = 42";
					}
					else
						context = " OR cod_situacao = 42";

					sqlFiltroSituacao += context;
				}
				else {
					if(sqlFiltroSituacao.indexOf(" AND cod_situacao = 42") != -1)
						context = " AND cod_situacao = 42";
					else if(sqlFiltroSituacao.indexOf(" OR cod_situacao = 42") != -1)
						context = " OR cod_situacao = 42";

					sqlFiltroSituacao = sqlFiltroSituacao.replace(context, "");

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						sqlFiltroSituacao = sqlFiltroSituacao.replace("OR", "AND");
					}

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						hasAndSqlQuery = false;
					}
				}

				context = "";
				if($(chk_paralizadas)[0].checked) {
					if(!hasAndSqlQuery){
						hasAndSqlQuery = true;
						context = " AND cod_situacao = 43";
					}
					else
						context = " OR cod_situacao = 43";

					sqlFiltroSituacao += context;
				}
				else {
					if(sqlFiltroSituacao.indexOf(" AND cod_situacao = 43") != -1)
						context = " AND cod_situacao = 43";
					else if(sqlFiltroSituacao.indexOf(" OR cod_situacao = 43") != -1)
						context = " OR cod_situacao = 43";

					sqlFiltroSituacao = sqlFiltroSituacao.replace(context, "");

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						sqlFiltroSituacao = sqlFiltroSituacao.replace("OR", "AND");
					}

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						hasAndSqlQuery = false;
					}
				}

				context = "";
				if($(chk_atendimento_programado)[0].checked) {
					if(!hasAndSqlQuery){
						hasAndSqlQuery = true;
						context = " AND cod_situacao = 44";
					}
					else
						context = " OR cod_situacao = 44";

					sqlFiltroSituacao += context;
				}
				else {
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

				context = "";
				if($(chk_nao_atendida)[0].checked) {
					if(!hasAndSqlQuery){
						hasAndSqlQuery = true;
						context = " AND cod_situacao = 41";
					}
					else
						context = " OR cod_situacao = 41";

					sqlFiltroSituacao += context;
				}
				else {
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

				context = "";
				if($(chk_atendimento_potencial)[0].checked) {
					if(!hasAndSqlQuery){
						hasAndSqlQuery = true;
						context = " AND cod_situacao = 45";
					}
					else
						context = " OR cod_situacao = 45";

					sqlFiltroSituacao += context;
				}
				else {
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

				if(sqlFiltroSituacao.indexOf("AND") != -1)
					loadEmpreendimentos(sqlFiltroSituacao);
				else
					$(".alert-situacao").removeClass("hide");
			});

			$.each($("input[type=checkbox]"), function(i,item){
				item.checked = true;
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
					<li><a href="#" class="expand"><i class="fa fa-expand"></i>&nbsp; Tela Cheia</a></li>
					<li><a href="<%= MM_Logout %>" class="sign-out"><i class="fa fa-sign-out"></i> Sair do Sistema</a></li>
				</ul>
			</div>
		</div>
	</nav>

	<div class="container">
		<div class="row">
			<div class="col-xs-4 sidebar">
				<form id="form-search" class="form" role="form">
					<div class="form-group find">
						<label class="control-label">
							Pesquisa de Municípios:
							<br/>
							<div class="alfa-links"></div>
						</label>
						<input class="form-control" id="txt-municipio" placeholder="Digite o nome do muncípio"></input>
					</div>

					<div class="form-group">
						<label class="control-label sr-only">buscar</label>
						<button type="submit" class="btn btn-block btn-primary"><i class="fa fa-search"></i> Buscar</button>
					</div>
				</form>

				<div class="form-group result hide">
					<label class="control-label">Resultado da Pesquisa:</label>
					<div class="alert alert-danger hide"><i class="fa fa-warning"></i> Nenhum município encontrado</div>
					<table class="table table-municipios">
						<tbody></tbody>
					</table>
					<button type="button" id="btn-clear" class="btn btn-danger btn-sm btn-block hide"><i class="fa fa-trash-o"></i> Limpar pesquisa</button>
				</div>

				<div class="form-group">
					<label class="control-label">Relatórios Consolidados:</label>
				</div>

				<div class="form-group">
					<label class="control-label sr-only"></label>
					<a href="resumo-situacao.asp?rep_universo_atendimento_programa=sim" class="btn btn-info btn-block btn-sm btn-text-left"><i class="fa fa-file-text-o"></i> Universo de Atendimento</a>
				</div>

				<div class="form-group">
					<label class="control-label sr-only"></label>
					<a href="resumo-situacao.asp?rep_universo_programa=sim" class="btn btn-info btn-block btn-sm btn-text-left"><i class="fa fa-file-text-o"></i> Universo do Programa</a>
				</div>

				<div class="form-group">
					<label class="control-label">Filtro por Situação:</label>

					<div class="checkbox">
						<label>
							<input id="chk-concluidas" type="checkbox"><img src="img/point_light_green.png"> Localidades com obras concluídas
						</label>
						<i class="fa fa-spinner fa-spin pull-right"></i>
						<button class="btn btn-xs btn-info pull-right hide btn-details btn-show-obras-concluidas">xxx Obra(s)</button>
					</div>

					<div class="checkbox">
						<label>
							<input id="chk-em-andamento" type="checkbox">
							<span style="padding-left: 10px;">Localidades com obras em andamento:</span>
						</label>
						<i class="fa fa-spinner fa-spin pull-right"></i>
						<button class="btn btn-xs btn-info pull-right hide btn-details btn-show-obras-em-andamento">xxx Obra(s)</button>
					</div>

					<!--  -->
					<div class="checkbox">
						<label style="padding-left: 50px;">
							<input id="chk-em-atendimento" type="checkbox"><img src="img/point_dark_green.png"> Obras em atendimento
						</label>
						<i class="fa fa-spinner fa-spin pull-right"></i>
						<button class="btn btn-xs btn-info pull-right hide btn-details btn-show-obras-em-atendimento">xxx Obra(s)</button>
					</div>

					<div class="checkbox">
						<label style="padding-left: 50px;">
							<input id="chk-finalizada-nao-operando" type="checkbox"><img src="img/point_orange.png"> Obras finalizadas e não operando
						</label>
						<i class="fa fa-spinner fa-spin pull-right"></i>
						<button class="btn btn-xs btn-info pull-right hide btn-details btn-show-obras-finalizada-nao-operando">xxx Obra(s)</button>
					</div>

					<div class="checkbox">
						<label style="padding-left: 50px;">
							<input id="chk-paralizadas" type="checkbox"><img src="img/point_purple.png"> Obras paralizadas
						</label>
						<i class="fa fa-spinner fa-spin pull-right"></i>
						<button class="btn btn-xs btn-info pull-right hide btn-details btn-show-obras-paralizadas">xxx Obra(s)</button>
					</div>
					<!--  -->

					<div class="checkbox">
						<label>
							<input id="chk-atendimento-programado" type="checkbox"><img src="img/point_light_blue.png"> Localidades com atendimento programado
						</label>
						<i class="fa fa-spinner fa-spin pull-right"></i>
						<button class="btn btn-xs btn-info pull-right hide btn-details btn-show-obras-atendimento-programado">xxx Obra(s)</button>
					</div>

					<div class="checkbox">
						<label>
							<input id="chk-nao-atendida" type="checkbox"><img src="img/point_red.png"> Localidades não atendidas
						</label>
						<i class="fa fa-spinner fa-spin pull-right"></i>
						<button class="btn btn-xs btn-info pull-right hide btn-details btn-show-obras-nao-atendidas">xxx Obra(s)</button>
					</div>

					<div class="checkbox">
						<label>
							<input id="chk-atendimento-potencial" type="checkbox"><img src="img/point_dark_blue.png"> Localidades com atendimento potencial
						</label>
						<i class="fa fa-spinner fa-spin pull-right"></i>
						<button class="btn btn-xs btn-info pull-right hide btn-details btn-show-obras-atendimento-potencial">xxx Obra(s)</button>
					</div>
				</div>

				<div class="form-group">
					<label class="control-label sr-only"></label>
					<button id="btn-load-situacoes" class="btn btn-primary btn-block"><i class="fa fa-filter"></i> Filtrar situações selecionadas</button>
				</div>

				<div class="form-group">
					<label class="control-label sr-only"></label>
					<div class="alert alert-warning alert-situacao hide"><i class="fa fa-warning"></i> Você deve selecionar ao menos uma situação!</div>
				</div>
			</div>

			<div class="col-xs-8 mapa">
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

	<div class="modal fade" id="modalQtdObras" tabindex="-1" role="dialog" aria-labelledby="modalQtdObrasLabel" aria-hidden="true">
		<div class="modal-dialog">
			<div class="modal-content">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					<h4 class="modal-title" id="modalQtdObrasLabel"></h4>
				</div>
				<div class="modal-body">
					<div class="row">
						<div class="col-xs-12">
							<div class="table-responsive">
								<table class="table table-localidades table-bordered table-condensed table-striped table-hover">
									<thead>
										<tr class="info">
											<th>Município</th>
											<th>Localidade</th>
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
					<button type="button" class="btn btn-default" data-dismiss="modal">Fechar</button>
				</div>
			</div>
		</div>
	</div>
</body>
</html>