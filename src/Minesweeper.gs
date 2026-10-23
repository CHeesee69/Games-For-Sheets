function startMinesweeper() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const gameRange = sheet.getRange("B2:K11");
  gameRange.clear();
  
  gameRange.setBackground('#4CAF50')
           .setBorder(true, true, true, true, true, true, '#388E3C', SpreadsheetApp.BorderStyle.SOLID)
           .setFontWeight('bold')
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle')
           .setValue('');
  
  for (let i = 2; i <= 11; i++) {
    sheet.setColumnWidth(i, 40);
    sheet.setRowHeight(i, 40);
  }
  
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('GAME_STARTED', 'false');
  scriptProperties.setProperty('GAME_OVER', 'false');
}

function generateMinesAndNumbers(firstR, firstC) {
  const SIZE = 10;
  const NUM_MINES = 15;
  let logicGrid = Array.from({ length: SIZE }, () => Array(SIZE).fill(0));
  
  let minesPlanted = 0;
  while (minesPlanted < NUM_MINES) {
    let r = Math.floor(Math.random() * SIZE);
    let c = Math.floor(Math.random() * SIZE);
    
    let isReserved = false;
    for (let i = -1; i <= 1; i++) {
      for (let j = -1; j <= 1; j++) {
        if (r === firstR + i && c === firstC + j) {
          isReserved = true;
        }
      }
    }
    
    if (logicGrid[r][c] !== 'M' && !isReserved) {
      logicGrid[r][c] = 'M';
      minesPlanted++;
    }
  }
  
  for (let r = 0; r < SIZE; r++) {
    for (let c = 0; c < SIZE; c++) {
      if (logicGrid[r][c] === 'M') continue;
      let count = 0;
      for (let i = -1; i <= 1; i++) {
        for (let j = -1; j <= 1; j++) {
          if (logicGrid[r + i] && logicGrid[r + i][c + j] === 'M') {
            count++;
          }
        }
      }
      logicGrid[r][c] = count;
    }
  }
  
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('LOGIC_GRID', JSON.stringify(logicGrid));
  scriptProperties.setProperty('GAME_STARTED', 'true');
  return logicGrid;
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const input = (e.value || '').trim().toUpperCase();
  
  const startRow = 2;
  const endRow = 11;
  const startCol = 2;
  const endCol = 11;
  
  const editRow = range.getRow();
  const editCol = range.getColumn();
  
  if (editRow < startRow || editRow > endRow || editCol < startCol || editCol > endCol) return;
  
  const scriptProperties = PropertiesService.getScriptProperties();
  if (scriptProperties.getProperty('GAME_OVER') === 'true') {
    range.setValue('');
    return;
  }
  
  const r = editRow - startRow;
  const c = editCol - startCol;
  const SIZE = 10;
  
  if (input === 'F') {
    range.setValue('');
    
    let currentVal = range.getValue();
    let currentBg = range.getBackground();
    if (currentBg === '#cccccc' || currentBg === '#CCCCCC') return;
    
    if (currentVal === '🚩') {
      range.setValue('').setBackground('#4CAF50').setFontColor('#000000');
    } else {
      range.setValue('🚩').setBackground('#81C784').setFontColor('#D32F2F');
    }
  } 
  else if (input === 'D') {
    if (range.getValue() === '🚩') {
      SpreadsheetApp.getUi().alert('Cell is flagged. Type F to unflag it before digging.');
      range.setValue('🚩');
      return;
    }
    
    let logicGrid;
    if (scriptProperties.getProperty('GAME_STARTED') === 'false') {
      logicGrid = generateMinesAndNumbers(r, c);
    } else {
      logicGrid = JSON.parse(scriptProperties.getProperty('LOGIC_GRID'));
    }
    
    revealCell(sheet, logicGrid, r, c, SIZE, startRow, startCol);
    checkWinCondition(sheet, logicGrid, SIZE, startRow, startCol);
  } 
  else {
    const oldText = e.oldValue || '';
    range.setValue(oldText);
  }
}

function revealCell(sheet, logicGrid, r, c, SIZE, startRow, startCol) {
  let cell = sheet.getRange(r + startRow, c + startCol);
  let val = logicGrid[r][c];
  let bg = cell.getBackground();
  
  if (bg === '#cccccc' || bg === '#CCCCCC' || bg === '#ff4444' || bg === '#FF4444' || cell.getValue() === '🚩') return;
  
  if (val === 'M') {
    cell.setBackground('#FF4444').setValue('');
    PropertiesService.getScriptProperties().setProperty('GAME_OVER', 'true');
    revealAllMines(sheet, logicGrid, SIZE, startRow, startCol);
    SpreadsheetApp.getUi().alert('Game Over!');
    return;
  }
  
  cell.setBackground('#CCCCCC');
  
  if (val > 0) {
    cell.setValue(val);
    const colors = ['#1976D2', '#388E3C', '#D32F2F', '#7B1FA2', '#FF8F00', '#0097A7', '#000000', '#777777'];
    cell.setFontColor(colors[val - 1]);
  } else {
    cell.setValue('');
    for (let i = -1; i <= 1; i++) {
      for (let j = -1; j <= 1; j++) {
        let nr = r + i;
        let nc = c + j;
        if (nr >= 0 && nr < SIZE && nc >= 0 && nc < SIZE) {
          revealCell(sheet, logicGrid, nr, nc, SIZE, startRow, startCol);
        }
      }
    }
  }
}

function revealAllMines(sheet, logicGrid, SIZE, startRow, startCol) {
  for (let r = 0; r < SIZE; r++) {
    for (let c = 0; c < SIZE; c++) {
      if (logicGrid[r][c] === 'M') {
        sheet.getRange(r + startRow, c + startCol).setBackground('#FF4444').setValue('');
      }
    }
  }
}

function checkWinCondition(sheet, logicGrid, SIZE, startRow, startCol) {
  let unrevealedSafeCells = 0;
  for (let r = 0; r < SIZE; r++) {
    for (let c = 0; c < SIZE; c++) {
      let bg = sheet.getRange(r + startRow, c + startCol).getBackground();
      let isRevealed = (bg === '#cccccc' || bg === '#CCCCCC');
      if (logicGrid[r][c] !== 'M' && !isRevealed) {
        unrevealedSafeCells++;
      }
    }
  }
  
  if (unrevealedSafeCells === 0) {
    PropertiesService.getScriptProperties().setProperty('GAME_OVER', 'true');
    SpreadsheetApp.getUi().alert('You Win!');
  }
}