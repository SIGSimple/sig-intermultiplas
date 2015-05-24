var map;

function initMap(){
	// objeto de preferências e configurações do mapa do google maps
	var mapOptions = {
		streetViewControl: false,
		scrollWheel: false,
		zoom: 18,
		mapTypeId: google.maps.MapTypeId.SATELLITE
	};

	// inicializa o google maps and adiciona-o ao html
	map = new google.maps.Map(document.getElementById("map-canvas"), mapOptions);
	
	var LatLng = getLatitudeLongitude();
		LatLng = replaceAll("°", "", LatLng);
	
	var latlog 		= LatLng.split(",");
	var latitude 	= ((latlog[0]).indexOf("-") >= 0) ? -Math.abs(latlog[0]) : Math.abs(latlog[0]);
	var longitude 	= ((latlog[1]).indexOf("-") >= 0) ? -Math.abs(latlog[1]) : Math.abs(latlog[1]);
	
	var location = new google.maps.LatLng(latitude, longitude);

	map.setCenter(location);

	var marker = new google.maps.Marker({
		map: map,
		position: location
	});

	var content = "Teste de Exibição";
	
	var infowindow = new google.maps.InfoWindow();

	google.maps.event.addListener(marker,'click', (function(marker,content,infowindow){
		return function() {
			infowindow.setContent(content);
			infowindow.open(map,marker);
			addButtonZoomClickListener();
		};
	})(marker,content,infowindow));
}

google.maps.event.addDomListener(window, 'load', initMap);
google.maps.event.addDomListener(window, "resize", resizingMap());

function resizeMap() {
   if(typeof map =="undefined") return;
   setTimeout( function(){resizingMap();} , 400);
}

function resizingMap() {
   if(typeof map =="undefined") return;
   var center = map.getCenter();
   google.maps.event.trigger(map, "resize");
   map.setCenter(center); 
}