function ms_startGame() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const gameRange = sheet.getRange("B2:K11");
  
  gameRange.clearDataValidations()
           .clearContent()
           .clearFormat();
  
  gameRange.setBackground('#4CAF50')
           .setBorder(true, true, true, true, true, true, '#388E3C', SpreadsheetApp.BorderStyle.SOLID)
           .setFontWeight('bold')
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle')
           .setFontColor('#000000')
           .setValue('');
  
  for (let i = 2; i <= 11; i++) {
    sheet.setColumnWidth(i, 40);
    sheet.setRowHeight(i, 40);
  }

  const labelCell = sheet.getRange("M9");
  labelCell.setValue("Mines:")
           .setFontWeight("bold")
           .setHorizontalAlignment("center");

  const mineCell = sheet.getRange("M10");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['10', '15', '20', '25'], true)
    .setAllowInvalid(false)
    .build();

  mineCell.setDataValidation(rule)
          .setFontWeight("bold")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");

  if (!mineCell.getValue()) {
    mineCell.setValue(15);
  }

  let dataSheet = ss.getSheetByName("_MinesweeperData_MS");
  if (!dataSheet) {
    dataSheet = ss.insertSheet("_MinesweeperData_MS");
  }
  dataSheet.clear();
  dataSheet.getRange("A1").setValue("GAME_STARTED");
  dataSheet.getRange("B1").setValue("false");
  dataSheet.getRange("A2").setValue("GAME_OVER");
  dataSheet.getRange("B2").setValue("false");
  dataSheet.hideSheet();
}

function ms_buildGrid(dataSheet, startR, startC, numMines) {
  const SIZE = 10;
  let matrix = Array.from({ length: SIZE }, () => Array(SIZE).fill(0));
  
  let minesPlanted = 0;
  while (minesPlanted < numMines) {
    let r = Math.floor(Math.random() * SIZE);
    let c = Math.floor(Math.random() * SIZE);
    let isReserved = Math.abs(r - startR) <= 1 && Math.abs(c - startC) <= 1;
    
    if (matrix[r][c] !== 'M' && !isReserved) {
      matrix[r][c] = 'M';
      minesPlanted++;
    }
  }
  
  for (let r = 0; r < SIZE; r++) {
    for (let c = 0; c < SIZE; c++) {
      if (matrix[r][c] === 'M') continue;
      let count = 0;
      for (let i = -1; i <= 1; i++) {
        for (let j = -1; j <= 1; j++) {
          let nr = r + i;
          let nc = c + j;
          if (nr >= 0 && nr < SIZE && nc >= 0 && nc < SIZE) {
            if (matrix[nr][nc] === 'M') count++;
          }
        }
      }
      matrix[r][c] = count;
    }
  }
  
  dataSheet.getRange(4, 1, SIZE, SIZE).setValues(matrix);
  dataSheet.getRange("B1").setValue("true");
  return matrix;
}

function handleMinesweeperEdit(e, editRow, editCol) {
  const range = e.range;
  const sheet = range.getSheet();
  const input = (e.value || '').toString().trim().toUpperCase();
  const startRow = 2;
  const startCol = 2;
  const SIZE = 10;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName("_MinesweeperData_MS");
  
  if (!dataSheet || dataSheet.getRange("B2").getValue().toString() === 'true') {
    range.setValue('');
    return;
  }
  
  const r = editRow - startRow;
  const c = editCol - startCol;
  
  if (input === 'F') {
    let currentVal = range.getValue();
    if (currentVal === '🚩') {
      range.setValue('').setBackground('#4CAF50').setFontColor('#000000');
    } else {
      range.setValue('🚩').setBackground('#81C784').setFontColor('#D32F2F');
    }
  } 
  else if (input === 'D') {
    if (range.getValue() === '🚩') {
      range.setValue('🚩');
      return;
    }
    
    let isStarted = dataSheet.getRange("B1").getValue().toString();
    
    let numMines = Number(sheet.getRange("M10").getValue());
    if (isNaN(numMines) || numMines <= 0) {
      numMines = 15;
    }

    let matrix = (isStarted !== 'true') 
      ? ms_buildGrid(dataSheet, r, c, numMines) 
      : dataSheet.getRange(4, 1, SIZE, SIZE).getValues();
    
    ms_openTile(sheet, dataSheet, matrix, r, c, SIZE, startRow, startCol);
  } 
  else {
    range.setValue('');
  }
}

function ms_openTile(sheet, dataSheet, matrix, r, c, SIZE, startRow, startCol) {
  let cell = sheet.getRange(r + startRow, c + startCol);
  let val = matrix[r][c];
  let bg = cell.getBackground().toLowerCase();
  
  if (bg === '#cccccc' || bg === '#ff4444' || cell.getValue() === '🚩') return;
  
  if (val === 'M' || val === 'm') {
    cell.setBackground('#FF4444');
    dataSheet.getRange("B2").setValue("true");
    for (let i = 0; i < SIZE; i++) {
      for (let j = 0; j < SIZE; j++) {
        if (matrix[i][j] === 'M' || matrix[i][j] === 'm') {
          sheet.getRange(i + startRow, j + startCol).setBackground('#FF4444');
        }
      }
    }
    return;
  }
  
  cell.setBackground('#CCCCCC');
  let numVal = Number(val);
  
  if (numVal > 0) {
    cell.setValue(numVal);
    const colors = ['#1976D2', '#388E3C', '#D32F2F', '#7B1FA2', '#FF8F00', '#0097A7', '#000000', '#777777'];
    cell.setFontColor(colors[numVal - 1]);
  } else {
    cell.setValue('');
    for (let i = -1; i <= 1; i++) {
      for (let j = -1; j <= 1; j++) {
        let nr = r + i;
        let nc = c + j;
        if (nr >= 0 && nr < SIZE && nc >= 0 && nc < SIZE) {
          ms_openTile(sheet, dataSheet, matrix, nr, nc, SIZE, startRow, startCol);
        }
      }
    }
  }
}

function startSudoku() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange("B2:J10");
  
  var difficulty = sheet.getRange("N4").getValue().toString().toLowerCase().trim();
  var clues = 42;
  if (difficulty === "medium") {
    clues = 32;
  } else if (difficulty === "hard") {
    clues = 24;
  }
  
  range.clearContent();
  range.setFontWeight("normal");
  range.setFontColor("#000000"); 
  range.setBackground("#FFFFFF");  
  
  let solutionBoard = createSolvedBoard();
  if (!solutionBoard) return;
  
  let puzzleBoard = carvePuzzle(solutionBoard, clues);
  for (let r = 0; r < 9; r++) {
    for (let c = 0; c < 9; c++) {
      let cell = range.getCell(r + 1, c + 1);
      if (puzzleBoard[r][c] !== 0) {
        cell.setValue(puzzleBoard[r][c]);
        cell.setFontWeight("bold");
        cell.setFontColor("#444444"); 
        cell.setBackground("#F1F3F4"); 
      } else {
        cell.setValue(""); 
      }
    }
  }
  
  PropertiesService.getScriptProperties().setProperty('sudokuSolution', JSON.stringify(solutionBoard));
}

function createSolvedBoard() {
  let grid = Array.from({ length: 9 }, () => Array(9).fill(0));
  if (fillBoard(grid)) return grid;
  return null;
}

function fillBoard(grid) {
  for (let row = 0; row < 9; row++) {
    for (let col = 0; col < 9; col++) {
      if (grid[row][col] === 0) {
        let nums = [1, 2, 3, 4, 5, 6, 7, 8, 9];
        for (let i = nums.length - 1; i > 0; i--) {
          const j = Math.floor(Math.random() * (i + 1));
          [nums[i], nums[j]] = [nums[j], nums[i]];
        }
        for (let num of nums) {
          if (isValidPlacement(grid, row, col, num)) {
            grid[row][col] = num;
            if (fillBoard(grid)) return true;
            grid[row][col] = 0; 
          }
        }
        return false; 
      }
    }
  }
  return true;
}

function isValidPlacement(grid, row, col, num) {
  for (let x = 0; x < 9; x++) {
    if (grid[row][x] === num || grid[x][col] === num) return false;
    
    let boxRow = 3 * Math.floor(row / 3) + Math.floor(x / 3);
    let boxCol = 3 * Math.floor(col / 3) + (x % 3);
    if (grid[boxRow][boxCol] === num) return false;
  }
  return true;
}

function carvePuzzle(board, clues) {
  let puzzle = board.map(r => [...r]);
  let cellsToRemove = 81 - clues;
  while (cellsToRemove > 0) {
    let r = Math.floor(Math.random() * 9);
    let c = Math.floor(Math.random() * 9);
    if (puzzle[r][c] !== 0) {
      puzzle[r][c] = 0;
      cellsToRemove--;
    }
  }
  return puzzle;
}

function handleSudokuEdit(e) {
  var range = e.range;
  var val = e.value;
  if (!val) return; 
  
  if (isNaN(val) || val < 1 || val > 9 || val.toString().includes('.')) {
    range.setValue("");
  }
}

function onEdit(e) {
  if (!e || !e.range) return;
  
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  
  if (row >= 2 && row <= 11 && col >= 2 && col <= 11) {
    handleMinesweeperEdit(e, row, col);
    return;
  }
  
  if (row >= 2 && row <= 10 && col >= 2 && col <= 10) {
    handleSudokuEdit(e);
    return;
  }
}