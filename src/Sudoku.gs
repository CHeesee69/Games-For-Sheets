function startSudoku() {
  const sht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sudoku");
  if (!sht) return;
  var rng = sht.getRange("B2:J10");
  var diffculty = sht.getRange("N4").getValue().toString().toLowerCase().trim();
  var cls = 42;
  if (diffculty === "medium") cls = 32;
  else if (diffculty === "hard") cls = 24;
  rng.clearContent();
  rng.setFontWeight("normal");
  rng.setFontColor("#000000");
  rng.setBackground("#FFFFFF");
  let solBoard = createSolvedBoard();
  if (!solBoard) return;
  let pzlBoard = carvePuzzle(solBoard, cls);
  for (let r = 0; r < 9; r++) {
    for (let c = 0; c < 9; c++) {
      let cl = rng.getCell(r + 1, c + 1);
      if (pzlBoard[r][c] !== 0) {
        cl.setValue(pzlBoard[r][c]).setFontWeight("bold").setFontColor("#444444").setBackground("#F1F3F4");
      } else {
        cl.setValue("");
      }
    }
  }
  PropertiesService.getScriptProperties().setProperty('sudokuSolution', JSON.stringify(solBoard));
}

function createSolvedBoard() {
  let grd = Array.from({ length: 9 }, () => Array(9).fill(0));
  if (fillBoard(grd)) return grd;
  return null;
}

function fillBoard(grd) {
  for (let row = 0; row < 9; row++) {
    for (let col = 0; col < 9; col++) {
      if (grd[row][col] === 0) {
        let nms = [1, 2, 3, 4, 5, 6, 7, 8, 9];
        for (let i = nms.length - 1; i > 0; i--) {
          const j = Math.floor(Math.random() * (i + 1));
          [nms[i], nms[j]] = [nms[j], nms[i]];
        }
        for (let num of nms) {
          if (isValidPlacement(grd, row, col, num)) {
            grd[row][col] = num;
            if (fillBoard(grd)) return true;
            grd[row][col] = 0;
          }
        }
        return false;
      }
    }
  }
  return true;
}

function isValidPlacement(grd, row, col, num) {
  for (let x = 0; x < 9; x++) {
    if (grd[row][x] === num || grd[x][col] === num) return false;
    let bxRow = 3 * Math.floor(row / 3) + Math.floor(x / 3);
    let bxCol = 3 * Math.floor(col / 3) + (x % 3);
    if (grd[bxRow][bxCol] === num) return false;
  }
  return true;
}

function carvePuzzle(brd, cls) {
  let pzl = brd.map(r => [...r]);
  let clsToRemve = 81 - cls;
  while (clsToRemve > 0) {
    let r = Math.floor(Math.random() * 9);
    let c = Math.floor(Math.random() * 9);
    if (pzl[r][c] !== 0) {
      pzl[r][c] = 0;
      clsToRemve--;
    }
  }
  return pzl;
}
