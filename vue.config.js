const fs = require('fs');
const path = require('path');
module.exports = {
	devServer: {
		https: {
			key: fs.readFileSync(path.join(__dirname, 'localhost+1-key.pem')),
			cert: fs.readFileSync(path.join(__dirname, 'localhost+1.pem')),
		},
		port: 8080,
	},
};
