const express = require('express');
const bodyParser = require('body-parser');
const updateCoordinates = require('./updateCoordinates');

module.exports.initExcel = function () {
	updateCoordinates.start();
};

module.exports.init = function () {
	let app = express();

	this.initExcel();

	return app;
};
