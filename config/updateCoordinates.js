const excelConfig = require('./excelConfig');
const Excel = require('exceljs'); // https://github.com/exceljs/exceljs
const fs = require('fs');

// ******
// Observações Janeiro 2022
// Novo batch de arquivos de Caio Paoli Janeiro 2022
// Esses arquivos possuem um novo layout em relação ao de October, por isso repliquei esse arquivo com as diferenças
// Parece que voltaram pra configuração original que está no arquivo excel.js
// Mesmo assim vou deixar aqui registrado as mudanças. Copiei do excelOctober e fui ajustando (ver onde tem comentado)
// OBS: Lembrar de abrir cada arquivo input e clicar em "Habilitar Modificações" e SALVAR
// ******
module.exports.start = async () => {
	console.log('---------------------- Início da Leitura dos Arquivos ----------------------');

	// Configuração das pastas:
	let cfg = {
		folderInput: excelConfig.config.path.folderInput,
		folderOutput: excelConfig.config.path.folderOutput,
		fileNameOutput: excelConfig.config.path.fileNameOutput,
	};
	let folders = fs.readdirSync(cfg.folderInput);
	let numeroEmpresas = 0;
	let numeroAtendimentos = 0;
	let numeroColaboradores = 0;
	let colecaoFichas = []; // Array final

	// Leitura para cada pasta (Empresa)
	for (let folder of folders) {
		numeroEmpresas++;
		let EMPRESA = folder;
		let numeroColaboradoresEmpresa = 0;
		let numeroAtendimentosEmpresa = 0;
		let folderPath = `${cfg.folderInput}\\${folder}`;
		let files = fs.readdirSync(folderPath);

		// Leitura para cada Arquivo (filial)
		for (let file of files) {
			let excelPathFile = `${folderPath}\\${file}`;
			console.log(`-> lendo   ${excelPathFile} `);

			const workbook = new Excel.Workbook();
			await workbook.xlsx.readFile(excelPathFile);
			// Abertura do arquivo excel, leitura para cada planilha:
			console.log({ planilhas: workbook.worksheets.length });

			workbook.eachSheet((worksheet, sheetId) => {
				let ficha = buildFicha();
				numeroColaboradores++;
				numeroColaboradoresEmpresa++;
				ficha.Colaborador.EMPRESA = EMPRESA;
				ficha.Colaborador.ARQUIVO = file;
				let atendSETUP = {
					lendoAtend: false,
					linha: 0,
					procurarOSSEO: 0,
				}; // Configuração que ajudará a determinar que estamos na fase de leitura de um atendiment (para via óssea)
				let atend = buildAtendimento();
				// Leitura para cada linha do arquivo
				// 1. Na coluna "A", ao encontrar "Ficha :" iniciará um novo colaborador (adiciona o colaborador anterior na array final)
				// 2. Na coluna "A", ao encontrar uma data: iniciará um atendimento para o colaborador (adiciona o atendimento anterior na array de atendimentos do colaborador)
				worksheet.eachRow((row, rowNumber) => {
					numeroAtendimentos++;
					numeroAtendimentosEmpresa++;

					let cellValue = readCellString(row, 1);
					if (atendSETUP.lendoAtend == true) {
						if (atendSETUP.procurarOSSEO == rowNumber) {
							atend.DIROSS500 = readCellString(row, 4); // January: Coluna D
							atend.DIROSS1 = readCellString(row, 5); // January: Coluna E
							atend.DIROSS2 = readCellString(row, 6); // January: Coluna F
							atend.DIROSS3 = readCellString(row, 7); // January: Coluna G
							atend.DIROSS4 = readCellString(row, 8); // January: Coluna H
							atend.DIROSS6 = readCellString(row, 9); // January: Coluna I

							atend.ESQOSS500 = readCellString(row, 13); // January: Coluna M
							atend.ESQOSS1 = readCellString(row, 14); // January: Coluna N
							atend.ESQOSS2 = readCellString(row, 15); // January: Coluna O
							atend.ESQOSS3 = readCellString(row, 16); // January: Coluna P
							atend.ESQOSS4 = readCellString(row, 17); // January: Coluna Q
							atend.ESQOSS6 = readCellString(row, 18); // January: Coluna R
						}
					}
					if (cellValue == null) return;
					if (cellValue.includes(`Ficha :`)) {
						// Armazena o colaborador anterior:
						colecaoFichas.push(ficha);
						ficha = buildFicha();
						numeroColaboradoresEmpresa++;
						numeroColaboradores++;
						ficha.Colaborador.EMPRESA = EMPRESA;
						ficha.Colaborador.ARQUIVO = file;

						ficha.Colaborador.Ficha = readCellString(row, 2);
					}
					if (cellValue.includes(`Dt. Nascimento :`)) {
						ficha.Colaborador.DtNascimento = readCellRaw(row, 3); // January: Coluna C
						ficha.Colaborador.IdadeAnos = readCellString(row, 6); // January: Coluna D
						ficha.Colaborador.IdadeMeses = readCellString(row, 8); // January: Coluna H
					}
					if (cellValue.includes(`Nome :`)) ficha.Colaborador.Nome = readCellString(row, 2);
					if (cellValue.includes(`Dt. Admissão :`)) ficha.Colaborador.DtAdmissao = readCellRaw(row, 3); // January: Coluna C
					if (cellValue.includes(`Cargo :`)) ficha.Colaborador.Cargo = readCellString(row, 2);
					if (cellValue.includes(`Local :`)) ficha.Colaborador.Local = readCellString(row, 2);

					// Verifica se temos uma data de atendimento
					if (readCellRaw(row, 1) instanceof Date) {
						atendSETUP.lendoAtend = true;
						atendSETUP.linha = rowNumber;
						atendSETUP.procurarOSSEO = rowNumber + 3;

						//atend = buildAtendimento();

						atend.Data = readCellRaw(row, 1);
						atend.DIRAER500 = readCellString(row, 4); // January: Coluna D
						atend.DIRAER1 = readCellString(row, 5); // January: Coluna E
						atend.DIRAER2 = readCellString(row, 6); // January: Coluna F
						atend.DIRAER3 = readCellString(row, 7); // January: Coluna G
						atend.DIRAER4 = readCellString(row, 8); // January: Coluna H
						atend.DIRAER6 = readCellString(row, 9); // January: Coluna I
						atend.DIRAER8 = readCellString(row, 10); // January: Coluna J

						atend.ESQAER500 = readCellString(row, 13); // January: Coluna M
						atend.ESQAER1 = readCellString(row, 14); // January: Coluna N
						atend.ESQAER2 = readCellString(row, 15); // January: Coluna O
						atend.ESQAER3 = readCellString(row, 16); // January: Coluna P
						atend.ESQAER4 = readCellString(row, 17); // January: Coluna Q
						atend.ESQAER6 = readCellString(row, 18); // January: Coluna R
						atend.ESQAER8 = readCellString(row, 19); // January: Coluna S
					}

					if (cellValue.includes(`Portaria 19`)) {
						atend.DIRPortaria19 = readCellString(row, 3); // January: Coluna C
						atend.ESQPortaria19 = readCellString(row, 12); // January: Coluna L
						ficha.Atendimentos.push(atend);
						atend = buildAtendimento();
						atendSETUP = {
							lendoAtend: false,
							linha: 0,
							procurarOSSEO: 0,
						};
					}
				});
				// Adiciona o último colaborador na array final
				colecaoFichas.push(ficha);
			});
			console.log(
				`<- fim     ${file}     ${numeroColaboradoresEmpresa} colaboradores      ${numeroAtendimentosEmpresa} atendimentos`,
			);
			console.log('');
		}

		// Iniciando o arquivo de saída com os dados da Empresa:
		let workbookOutput = new Excel.Workbook();
		let sheet = workbookOutput.addWorksheet('audios');
		// Adiciona a primeira linha (cabeçalho)
		sheet.addRow([
			'EMPRESA',
			'Arquivo',
			'FICHA',
			'NOMECOLAB',
			'DTNASCIMENTO',
			'IDADEANOS',
			'IDADEMESES',
			'DTADMISSAO',
			'CARGO',
			'SETOR',
			'DATAATENDIMENTO',
			'DIRAER500',
			'DIRAER1000',
			'DIRAER2000',
			'DIRAER3000',
			'DIRAER4000',
			'DIRAER6000',
			'DIRAER8000',
			'ESQAER500',
			'ESQAER1000',
			'ESQAER2000',
			'ESQAER3000',
			'ESQAER4000',
			'ESQAER6000',
			'ESQAER8000',
			'DIROSS500',
			'DIROSS1000',
			'DIROSS2000',
			'DIROSS3000',
			'DIROSS4000',
			'DIROSS6000',
			'ESQOSS500',
			'ESQOSS1000',
			'ESQOSS2000',
			'ESQOSS3000',
			'ESQOSS4000',
			'ESQOSS6000',
			'DIRPortaria19',
			'ESQPortaria19',
		]);

		// Leitura para cada colaborador:
		colecaoFichas.forEach((ficha) => {
			colabArray = [
				ficha.Colaborador.EMPRESA,
				ficha.Colaborador.ARQUIVO,
				ficha.Colaborador.Ficha,
				ficha.Colaborador.Nome,
				ficha.Colaborador.DtNascimento,
				ficha.Colaborador.IdadeAnos,
				ficha.Colaborador.IdadeMeses,
				ficha.Colaborador.DtAdmissao,
				ficha.Colaborador.Cargo,
				ficha.Colaborador.Local,
			];

			// Para cada atendimento do colaborador, adiciona uma linha com todas as informações do colaborador e do atendimento:
			ficha.Atendimentos.forEach((atend) => {
				sheet.addRow([
					...colabArray,
					atend.Data,
					atend.DIRAER500,
					atend.DIRAER1,
					atend.DIRAER2,
					atend.DIRAER3,
					atend.DIRAER4,
					atend.DIRAER6,
					atend.DIRAER8,
					atend.ESQAER500,
					atend.ESQAER1,
					atend.ESQAER2,
					atend.ESQAER3,
					atend.ESQAER4,
					atend.ESQAER6,
					atend.ESQAER8,
					atend.DIROSS500 == 'NT' ? '' : atend.DIROSS500,
					atend.DIROSS1 == 'NT' ? '' : atend.DIROSS1,
					atend.DIROSS2 == 'NT' ? '' : atend.DIROSS2,
					atend.DIROSS3 == 'NT' ? '' : atend.DIROSS3,
					atend.DIROSS4 == 'NT' ? '' : atend.DIROSS4,
					atend.DIROSS6 == 'NT' ? '' : atend.DIROSS6,
					atend.ESQOSS500 == 'NT' ? '' : atend.ESQOSS500,
					atend.ESQOSS1 == 'NT' ? '' : atend.ESQOSS1,
					atend.ESQOSS2 == 'NT' ? '' : atend.ESQOSS2,
					atend.ESQOSS3 == 'NT' ? '' : atend.ESQOSS3,
					atend.ESQOSS4 == 'NT' ? '' : atend.ESQOSS4,
					atend.ESQOSS6 == 'NT' ? '' : atend.ESQOSS6,
					atend.DIRPortaria19,
					atend.ESQPortaria19,
				]);
			});
		});

		// Salva o arquivo final da Empresa:
		await workbookOutput.xlsx.writeFile(cfg.fileNameOutput + '_' + EMPRESA + '.xlsx');

		// Reseta a coleção de colaboradores
		colecaoFichas = [];
	}

	console.log(
		`- Final        ${numeroEmpresas} Empresas      ${numeroColaboradores} Colaboradores(total)      ${numeroAtendimentos} Atendimentos(total)`,
	);
	console.log(`${cfg.fileNameOutput}`);
	console.log('');
	console.log('---------------------- Fim da Leitura dos Arquivos ----------------------');
};

// Leitura de uma célula, devolve uma string:
function readCellString(row, number) {
	let rowValue = row.values[number];

	if (rowValue == null) return null;

	return rowValue.toString();
}

// Leitura de uma célula, não converte o valor (para Datas):
function readCellRaw(row, number) {
	let rowValue = row.values[number];

	if (rowValue == null) return null;

	return rowValue;
}

// Inicia um novo objeto de colaborador:
function buildFicha() {
	return {
		Colaborador: {
			EMPRESA: '',
			ARQUIVO: '',
			Ficha: '',
			Nome: '',
			DtNascimento: '',
			IdadeAnos: '',
			IdadeMeses: '',
			DtAdmissao: '',
			Cargo: '',
			Local: '',
		},
		Atendimentos: [],
	};
}

// Inicia um novo objeto de atendimento:
function buildAtendimento() {
	return {
		Data: '',
		DIRAER500: '',
		DIRAER1: '',
		DIRAER2: '',
		DIRAER3: '',
		DIRAER4: '',
		DIRAER6: '',
		DIRAER8: '',
		ESQAER500: '',
		ESQAER1: '',
		ESQAER2: '',
		ESQAER3: '',
		ESQAER4: '',
		ESQAER6: '',
		ESQAER8: '',
		DIROSS500: '',
		DIROSS1: '',
		DIROSS2: '',
		DIROSS3: '',
		DIROSS4: '',
		DIROSS6: '',
		ESQOSS500: '',
		ESQOSS1: '',
		ESQOSS2: '',
		ESQOSS3: '',
		ESQOSS4: '',
		ESQOSS6: '',
		DIRPortaria19: '',
		ESQPortaria19: '',
	};
}
