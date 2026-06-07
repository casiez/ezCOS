// Gery Casiez - 2026

const fs = require('fs');
const carbone = require('carbone');

carbone.addFormatters({
	intMember : function (interne, intext) { 
	  if (interne === true && intext === "int") return '\u2611';
	  if (interne === false && intext === "int") return '\u2610';
	  if (interne === true && intext === "ext") return '\u2610';
	  if (interne === false && intext === "ext") return '\u2611';
	}
});

fs.readFile('membres.json', 'utf8', function (err, d) {
	if (err) throw err;
	data = JSON.parse(d);
	data.forEach(function(entry){ 
	  carbone.render('templates/Impartialite.xlsx', entry, function(err, result){
	    if (err) return console.log(err);
	    fs.writeFileSync('impartialite/' + entry['nomfichier'] + '.xlsx', result);
	  });
	});
});


