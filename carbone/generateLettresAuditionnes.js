// Gery Casiez - 2026

const fs = require('fs');
const carbone = require('carbone');

function convBR (d) { 
  if (d) {
    return d.replace(/\r?\n/g, '<w:br/>');
  }
  return "";
}

convBR.canInjectXML = true;

carbone.addFormatters({
  convBR
});

fs.readFile('candidats.json', 'utf8', function (err, d) {
	if (err) throw err;
	data = JSON.parse(d);

	data.forEach(function(entry){ 
	  carbone.render('templates/Convocation_candidats.docx', entry, function(err, result){
	    if (err) return console.log(err);
	    fs.writeFileSync('convocationsCandiats/' + entry['nomfichierCandidat'] + '.docx', result);
	  });
	});
});


