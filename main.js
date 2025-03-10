const { app, BrowserWindow, dialog } = require('electron');
const fs = require('fs');
const readline = require('readline');
const path = require('path');
const xlsx = require('xlsx');

function createWindow() {
  const win = new BrowserWindow({
    show: false,
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    }
  });

  dialog.showOpenDialog(win, {
    properties: ['openFile']
  }).then(result => {
    if (!result.canceled && result.filePaths.length > 0) {
      const filePath = result.filePaths[0];
      const fileName = path.basename(filePath);

      const fileStream = fs.createReadStream(filePath, 'utf8');
      const rl = readline.createInterface({
        input: fileStream,
        crlfDelay: Infinity
      });

      let currentPGM = null;
      let currentDDName = null;
      const output = [];

      rl.on('line', (line) => {
        line = line.trim();

        if (line.startsWith('//*')) return;

        const pgmMatch = line.match(/EXEC\s+PGM=(\w+)/);
        if (pgmMatch) {
          currentPGM = pgmMatch[1];
          currentDDName = null;
          return;
        }

        const ddLineMatch = line.match(/^(\w+)\s+DD\s+/);
        if (ddLineMatch) {
          currentDDName = ddLineMatch[1];
        }

        const dsnMatch = line.match(/DSN=([^,]+)/);
        if (dsnMatch) {
          const dsn = dsnMatch[1];

          const dispMatch = line.match(/DISP=([^,\s)]+)/) || line.match(/DISP=(\([^)]*\)),?/);
          const disp = dispMatch ? dispMatch[1] : '';

          output.push({
            FILENAME: fileName,
            PGM: currentPGM,
            PROC: '',
            'DD NAME': currentDDName,
            DSN: dsn,
            DISP: disp,
            'I/O': disp == 'SHR' ? 'Input' : '',
            'Created In': '',
          });
        }
      });

      rl.on('close', () => {
        console.log('Finished parsing file.\n');
        console.log(JSON.stringify(output, null, 2));

        const outputDir = path.join(__dirname, 'generated_files');
        if (!fs.existsSync(outputDir)) {
          fs.mkdirSync(outputDir);
        }

        const workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.json_to_sheet(output);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'PROC_DATA');

        const excelFileName = path.basename(fileName, path.extname(fileName)) + '.xlsx';
        const excelFilePath = path.join(outputDir, excelFileName);
        
        xlsx.writeFile(workbook, excelFilePath);

        console.log(`Excel file generated: ${excelFileName}`);

        win.close();
        app.quit();
      });
    }
  }).catch(err => {
    console.error('Error opening file dialog:', err);
  });
}

app.whenReady().then(createWindow);
