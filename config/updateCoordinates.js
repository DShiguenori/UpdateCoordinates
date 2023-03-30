const axios = require('axios');
const excelConfig = require('./excelConfig');
const Excel = require('exceljs'); // https://github.com/exceljs/exceljs
const fs = require('fs');
const { func } = require('edge-js');
require('dotenv').config();

module.exports.start = async () => {
	console.log('---------------------- Início da Leitura dos Arquivos ----------------------');

	// Folder configs:
	let cfg = {
		folderInput: excelConfig.config.path.folderInput,
		folderOutput: excelConfig.config.path.folderOutput,
	};
	let folderInput = `${cfg.folderInput}`;
	let folderOutput = `${cfg.folderOutput}`;
	let files = fs.readdirSync(folderInput);

	// Leitura para cada Arquivo (filial)
	for (let file of files) {
		let excelPathFile = `${folderInput}\\${file}`;
		console.log(`-> REDING...   ${excelPathFile} `);

		const excelInput = new Excel.Workbook();
		await excelInput.xlsx.readFile(excelPathFile);

		let workbookOutput = new Excel.Workbook();

		// Open the excel file, read each worksheet:
		excelInput.eachSheet((inSheet, sheetId) => {
			let outSheet = workbookOutput.addWorksheet('endereco_alunos_uptd');

			inSheet.eachRow(async (row, rowNumber) => {
				if (rowNumber == 1) {
					outSheet.addRow([
						readCellString(row, 1),
						readCellString(row, 2),
						readCellString(row, 3),
						readCellString(row, 4),
						readCellString(row, 5),
						readCellString(row, 6),
						'latitude',
						'longitude',
					]);
					return;
				}

				let completeAddress = getCompleteAddress(row);

				await sleep(1001); // Wait 1 second to not exceed the rate limit
				let coordinates = await getCoordinatesBySearch(completeAddress);
				if (coordinates == null) {
					let googleCoordinates = await getCoordinatesFromGoogle(completeAddress);
					if (googleCoordinates.latitude != null || googleCoordinates.longitude != null) {
						console.log('Coordinates obtained via Google Maps API');
						completeAddress.latitude = googleCoordinates.latitude;
						completeAddress.longitude = googleCoordinates.longitude;
					}
				} else {
					console.log('Coordinates obtained via OSM Nominatim API');
					completeAddress.latitude = coordinates.latitude;
					completeAddress.longitude = coordinates.longitude;
				}
				console.log(completeAddress.latitude + ' ' + completeAddress.longitude);

				outSheet.addRow([
					completeAddress.codigo_aluno,
					completeAddress.logradouro_endereco,
					completeAddress.bairro_endereco,
					completeAddress.cidade_endereco,
					completeAddress.uf_endereco,
					completeAddress.cep,
					completeAddress.latitude,
					completeAddress.longitude,
				]);
			});
		});
		await workbookOutput.xlsx.writeFile(`${folderOutput}\\out_${file}`);
		console.log(`FILE SAVED ${`${folderOutput}\\out_${file}`}`);
	}
	console.log('---------------------- END READING FOLDERS AND THEIR FILES ----------------------');
};

// Read a cell, outputs a string
function readCellString(row, number) {
	let rowValue = row.values[number];

	if (rowValue == null) return null;

	return rowValue.toString();
}
// Read a cell, does not convert the value (for Dates)
function readCellRaw(row, number) {
	let rowValue = row.values[number];

	if (rowValue == null) return null;

	return rowValue;
}

function sleep(ms) {
	return new Promise((resolve) => setTimeout(resolve, ms));
}

// Build the address object
function getCompleteAddress(row) {
	const address = {
		codigo_aluno: readCellString(row, 1),
		logradouro_endereco: readCellString(row, 2),
		bairro_endereco: readCellString(row, 3),
		cidade_endereco: readCellString(row, 4),
		uf_endereco: readCellString(row, 5),
		cep: readCellString(row, 6),
		latitude: '',
		longitude: '',
	};
	return address;
}

async function getCoordinatesBySearch(completeAddress) {
	const addressString = [completeAddress.cidade_endereco, completeAddress.logradouro_endereco].join(', ');

	const url = `https://nominatim.openstreetmap.org/search?format=json&limit=1&q=${encodeURIComponent(addressString)}`;
	console.log(url);
	try {
		const response = await axios.get(url, {
			headers: {
				'User-Agent': 'GetCoordinatesForStudents', // Replace 'YourAppName' with your actual app name
			},
		});
		const data = response.data;
		if (data && data.length > 0) {
			const coordinatesValues = {
				latitude: parseFloat(data[0].lat),
				longitude: parseFloat(data[0].lon),
			};
			return coordinatesValues;
		} else {
			return null;
		}
	} catch (error) {
		console.error(`Error fetching coordinates for ${addressString}: ${error.message}`);
		return { latitude: null, longitude: null };
	}
}

async function getCoordinatesFromGoogle(completeAddress) {
	const addressString = [completeAddress.cidade_endereco, completeAddress.logradouro_endereco].join(', ');

	const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(addressString)}&key=${
		process.env.GOOGLE_MAPS_API_KEY
	}`;
	console.log(url);
	try {
		const response = await axios.get(url);
		const data = response.data;
		if (data.status === 'OK') {
			const location = data.results[0].geometry.location;
			const coordinates = {
				latitude: location.lat,
				longitude: location.lng,
			};
			return coordinates;
		} else {
			console.log('Google Maps API failed to find coordinates');
			return { latitude: null, longitude: null };
		}
	} catch (error) {
		console.error(`Error fetching coordinates from Google Maps API for ${addressString}: ${error.message}`);
		return { latitude: null, longitude: null };
	}
}
