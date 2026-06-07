// Gery Casiez - 2020

const fs = require('fs');
const carbone = require('carbone');

carbone.addFormatters({
	convBR : function (d) { 
	      var res = d.replace(/\r?\n/g, '<w:br/>');
	      return res;
	}
});

carbone.formatters.convBR.canInjectXML = true;

fs.readFile('candidats.json', 'utf8', function (err, d) {
	if (err) throw err;
	data = JSON.parse(d);

	carbone.render('templates/PVrepartition.docx', data, function(err, result){
	  if (err) return console.log(err);
	  fs.writeFileSync('PVrepartitionRapporteurs.docx', result);
	});
});


