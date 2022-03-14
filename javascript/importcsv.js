// A CSV parser that complies with EXCEL import rules.
// The speed of the csv parser installed by npm is not good enough, so I create it.
// I'm new to javascript, so it may be unsightly, but please try it if you like.

class SpecialRapidService {
	// Special Rapid Service is a train that connects Kyoto and Osaka in a time comparable
	// to the Shinkansen and does not require a limited express fare. I have been indebted
	// to this train for a long time, so I chose the class name in honor of it.
	constructor(delim) {
		this.delim = (delim === undefined) ? ',' : delim;
	}

	parse(fs, callback, finish) {
		var record = [], quote = null;
		if (callback === undefined) callback = function(record) { };

		const rl = require('readline').createInterface({input: fs});

		rl.on('line', (line) => {
			var chunk = line.split(this.delim), grains;

			for (var i=0; i<chunk.length; i++) {
				if (quote === null) {
					if (chunk[i].charAt(0) !== '"') {
						record.push(chunk[i]);
						continue;
					}

					quote = '';
					grains = chunk[i].substr(1).split('"');
				}
				else
					grains = chunk[i].split('"');

				var result = false, limit = grains.length - 1;

				for (var j=1; ; j+=2) {
					if (j < limit) {
						if (grains[j] !== '') {
							result = true;
							while (++j<grains.length)
								grains[j] = '"' + grains[j];
							break;
						}
						grains[j] = '"';
					}
					else {
						if (j === limit) result = true;
						break;
					}
				}

				quote = quote + grains.join("")

				if (result) {
					record.push(quote);
					quote = null;
				}
				else
					quote = quote + (i + 1 === chunk.length ? '\n' : this.delim);
			}

			if (quote === null) {
				callback(record);
				record = [];
			}
		});

		rl.on('close', () => {
			if (quote !== null) {
				record.push(quote);
				callback(record);
			}

			if (finish !== undefined) finish();
		});
	}
}


var start = new Date();
rs = require("fs").createReadStream('./test.txt');

var srs = new SpecialRapidService();
srs.parse(rs,
	(record) => {	/* Callback for each record read */
		for (var i=0; i<record.length; i++) {
			console.log(i + " : " + record[i]);
		}
	},
	() => {			/* Callback at completed */
		console.log((new Date() - start) / 1000);
	});
