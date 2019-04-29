function getInputs(){
	var SEL = document.getElementById("SEL");
	var BOB = document.getElementById("BOB");
	var HUM = document.getElementById("HUM");
	var SEL_file = SEL.files[0];
	var BOB_file = BOB.files[0];
	var HUM_file = HUM.files[0];

	var string = SEL_file.name + BOB_file.name + HUM_file.name;

	alert(string);
}