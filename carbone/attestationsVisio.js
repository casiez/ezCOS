// Gery Casiez - 2026

const fs = require('fs');
const carbone = require('carbone');

fs.readFile('membres.json', 'utf8', function (err, d) {
	if (err) throw err;
	data = JSON.parse(d);
	data.forEach(function(entry){ 
	  carbone.render('templates/Attestation_visioconference_membre.docx', entry, function(err, result){
	    if (err) return console.log(err);
	    fs.writeFileSync('attestationsVisio/' + entry['nomfichier'] + '.docx', result);
	  });
	});
});


