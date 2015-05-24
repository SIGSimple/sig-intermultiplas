function capitalizeFirstLetter(string) {
	return string.charAt(0).toUpperCase() + string.slice(1);
}

function replaceAll(find, replace, str) {
	return str.replace(new RegExp(find, 'g'), replace);
}

function adjustNumLayout() {
	$.each($(".num"), function(i, item){
		$(item).val($.number($(item).val(), 0, ",", "."));
		$(item).text($.number($(item).text(), 0, ",", "."));
	});
}

function adjustVlrLayout() {
	$.each($(".vlr"), function(i, item){
		$(item).val($.number($(item).val(), 0, ",", "."));
		if($(item).text() != "")
			$(item).text("R$ " + $.number($(item).text(), 2, ",", "."));
	});
}

function adjustPrcLayout() {
	$.each($(".prc"), function(i,item){
		$(item).text($.number(( parseFloat($(item).text()) * 100 ), 2, ",", ".") + "%");
	});
}