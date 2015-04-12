$(function(){
	$("#form-upload").on("submit", function(e) {
		if(e.target.blob.value){
			$("#btnSubmit").attr("disabled","disabled").val("Aguarde...");
			return true;
		}
		else {
			alert("Você deve selecionar um arquivo!")
			return false;
		}
	});
});