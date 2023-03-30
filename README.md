# UpdateCoordinates

This mini-project is used to update the adresses with the lat-lon coordinates

### Order of execution:

1. server.js
2. app.js
3. express.js
4. updateCoordinates.js

### Configuring the folders:

You need to creat the folders for input and output and update the config file

Example:

```js
const config = {
	path: {
		folderInput: `C:\\Trabalho\\ECidade\\UpdateCoordinates\\excelIn`,
		folderOutput: `C:\\Trabalho\\ECidade\\UpdateCoordinates\\excelOut`,
	},
};

module.exports = {
	config,
};
```

### Reminder!!!

> Always open the excel file and Enable Editing
