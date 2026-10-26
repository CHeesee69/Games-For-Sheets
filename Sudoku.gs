function ms_startGame() {
  const sht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Minesweeper");
  if (!sht) return;
  const gmRange = sht.getRange("B2:K11");
  gmRange.clearDataValidations().clearContent().clearFormat();
  gmRange.setBackground('#4CAF50')
    .setBorder(true, true, true, true, true, true, '#388E3C', SpreadsheetApp.BorderStyle.SOLID)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontColor('#000000')
    .setValue('');
  for (let i = 2; i <= 11; i++) {
    sht.setColumnWidth(i, 40);
    sht.setRowHeight(i, 40);
  }
  const lblCell = sht.getRange("M9");
  lblCell.setValue("Mines:").setFontWeight("bold").setHorizontalAlignment("center");
  const mnCell = sht.getRange("M10");
  const rl = SpreadsheetApp.newDataValidation()
    .requireValueInList(['10', '15', '20', '25'], true)
    .setAllowInvalid(false)
    .build();
  mnCell.setDataValidation(rl).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  if (!mnCell.getValue()) {
    mnCell.setValue(15);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dtSheet = ss.getSheetByName("_MinesweeperData_MS");
  if (!dtSheet) {
    dtSheet = ss.insertSheet("_MinesweeperData_MS");
  }
  dtSheet.clear();
  dtSheet.getRange("A1").setValue("GAME_STARTED");
  dtSheet.getRange("B1").setValue("false");
  dtSheet.getRange("A2").setValue("GAME_OVER");
  dtSheet.getRange("B2").setValue("false");
  dtSheet.hideSheet();
}

function ms_buildGrid(dtSheet, strtR, strtC, nmMines) {
  const SZ = 10;
  let mtrx = Array.from({ length: SZ }, () => Array(SZ).fill(0));
  let mnsPlanted = 0;
  while (mnsPlanted < nmMines) {
    let r = Math.floor(Math.random() * SZ);
    let c = Math.floor(Math.random() * SZ);
    let isRservd = Math.abs(r - strtR) <= 1 && Math.abs(c - strtC) <= 1;
    if (mtrx[r][c] !== 'M' && !isRservd) {
      mtrx[r][c] = 'M';
      mnsPlanted++;
    }
  }
  for (let r = 0; r < SZ; r++) {
    for (let c = 0; c < SZ; c++) {
      if (mtrx[r][c] === 'M') continue;
      let cnt = 0;
      for (let i = -1; i <= 1; i++) {
        for (let j = -1; j <= 1; j++) {
          let nr = r + i;
          let nc = c + j;
          if (nr >= 0 && nr < SZ && nc >= 0 && nc < SZ) {
            if (mtrx[nr][nc] === 'M') cnt++;
          }
        }
      }
      mtrx[r][c] = cnt;
    }
  }
  dtSheet.getRange(4, 1, SZ, SZ).setValues(mtrx);
  dtSheet.getRange("B1").setValue("true");
  return mtrx;
}

function handleMinesweeperEdit(e) {
  const rng = e.range;
  const sht = rng.getSheet();
  const ss = sht.getParent();
  const edtRow = rng.getRow();
  const edtCol = rng.getColumn();
  const inpt = (e.value || '').toString().trim().toUpperCase();
  const strtRow = 2;
  const strtCol = 2;
  const SZ = 10;
  let dtSheet = ss.getSheetByName("_MinesweeperData_MS");

  if (!dtSheet || dtSheet.getRange("B2").getValue().toString() === 'true') {
    rng.setValue('');
    return;
  }

  const r = edtRow - strtRow;
  const c = edtCol - strtCol;

  if (inpt === 'F') {
    let currVal = rng.getValue();
    if (currVal === '🚩') {
      rng.setValue('').setBackground('#4CAF50').setFontColor('#000000');
    } else {
      rng.setValue('🚩').setBackground('#81C784').setFontColor('#D32F2F');
    }
  } else if (inpt === 'D') {
    rng.setValue('');

    if (rng.getValue() === '🚩') {
      rng.setValue('🚩');
      return;
    }
    let isStrted = dtSheet.getRange("B1").getValue().toString();
    let nmMines = Number(sht.getRange("M10").getValue());
    if (isNaN(nmMines) || nmMines <= 0) {
      nmMines = 15;
    }
    let mtrx = (isStrted !== 'true')
      ? ms_buildGrid(dtSheet, r, c, nmMines)
      : dtSheet.getRange(4, 1, SZ, SZ).getValues();

    let vsited = new Set();
    ms_openTile(sht, dtSheet, mtrx, r, c, SZ, strtRow, strtCol, vsited);
  } else {
    rng.setValue('');
  }
}

function ms_openTile(sht, dtSheet, mtrx, r, c, SZ, strtRow, strtCol, vsited) {
  let ky = r + ',' + c;
  if (vsited.has(ky)) return;
  vsited.add(ky);

  let cl = sht.getRange(r + strtRow, c + strtCol);
  let vl = mtrx[r][c];
  let bgColors = cl.getBackground().toLowerCase();
  
  if (bgColors === '#cccccc' || bgColors === '#ff4444' || cl.getValue() === '🚩') return;

  if (vl === 'M' || vl === 'm') {
    cl.setBackground('#FF4444');
    dtSheet.getRange("B2").setValue("true");
    let bgColrsArr = Array.from({ length: SZ }, () => Array(SZ).fill(null));
    const currBgs = sht.getRange(strtRow, strtCol, SZ, SZ).getBackgrounds();
    for (let i = 0; i < SZ; i++) {
      for (let j = 0; j < SZ; j++) {
        if (mtrx[i][j] === 'M' || mtrx[i][j] === 'm') {
          bgColrsArr[i][j] = '#FF4444';
        } else {
          bgColrsArr[i][j] = currBgs[i][j];
        }
      }
    }
    sht.getRange(strtRow, strtCol, SZ, SZ).setBackgrounds(bgColrsArr);
    return;
  }

  cl.setBackground('#CCCCCC');
  let nmVal = Number(vl);
  if (nmVal > 0) {
    cl.setValue(nmVal);
    const clrs = ['#1976D2', '#388E3C', '#D32F2F', '#7B1FA2', '#FF8F00', '#0097A7', '#000000', '#777777'];
    cl.setFontColor(clrs[nmVal - 1]);
  } else {
    cl.setValue('');
    for (let i = -1; i <= 1; i++) {
      for (let j = -1; j <= 1; j++) {
        let nr = r + i;
        let nc = c + j;
        if (nr >= 0 && nr < SZ && nc >= 0 && nc < SZ) {
          ms_openTile(sht, dtSheet, mtrx, nr, nc, SZ, strtRow, strtCol, vsited);
        }
      }
    }
  }
}
