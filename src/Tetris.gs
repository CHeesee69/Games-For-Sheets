function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const a1 = e.range.getA1Notation();
  if (sheetName === "Tetris") {
    if (a1 === "N8") addInputToQueue("LEFT");
    else if (a1 === "P8") addInputToQueue("RIGHT");
    else if (a1 === "Q8") addInputToQueue("ROTATE");
    return;
  }
  if (sheetName === "Minesweeper") {
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row >= 2 && row <= 11 && col >= 2 && col <= 11) {
      handleMinesweeperEdit(e);
    }
    return;
  }
}
const ROWS = 20;
const COLS = 10;
const TICK_SPEED_MS = 500;

const SHAPES = {
  'I': [[0,0,0,0],[1,1,1,1],[0,0,0,0],[0,0,0,0]],
  'O': [[1,1],[1,1]],
  'T': [[0,1,0],[1,1,1],[0,0,0]],
  'S': [[0,1,1],[1,1,0],[0,0,0]],
  'Z': [[1,1,0],[0,1,1],[0,0,0]],
  'J': [[1,0,0],[1,1,1],[0,0,0]],
  'L': [[0,0,1],[1,1,1],[0,0,0]]
};

function saveGameState(board, currentPiece, currentX, currentY, isPlaying) {
  var cache = CacheService.getScriptCache();
  var state = {
    board: board,
    currentPiece: currentPiece,
    currentX: currentX,
    currentY: currentY,
    isPlaying: isPlaying
  };
  cache.put('tetris_state', JSON.stringify(state), 300);
}

function loadGameState() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('tetris_state');
  if (cached) {
    return JSON.parse(cached);
  }
  return null;
}

function addInputToQueue(action) {
  var cache = CacheService.getScriptCache();
  cache.put('tetris_input', action, 10);
}

function popInputQueue() {
  var cache = CacheService.getScriptCache();
  var action = cache.get('tetris_input');
  if (action) {
    cache.remove('tetris_input');
    return action;
  }
  return null;
}

function checkCollision(nextX, nextY, piece, boardData) {
  for (var r = 0; r < piece.length; r++) {
    for (var c = 0; c < piece[r].length; c++) {
      if (piece[r][c] !== 0) {
        var targetX = nextX + c;
        var targetY = nextY + r;
        if (targetX < 0 || targetX >= COLS || targetY >= ROWS) return true;
        if (targetY >= 0 && boardData[targetY][targetX] !== 0) return true;
      }
    }
  }
  return false;
}

function rotateMatrix(matrix) {
  var n = matrix.length;
  var ret = [];
  for (var i = 0; i < n; i++) {
    ret.push(new Array(n).fill(0));
  }
  for (var i = 0; i < n; i++) {
    for (var j = 0; j < n; j++) {
      ret[j][n - 1 - i] = matrix[i][j];
    }
  }
  return ret;
}

function clearLines(boardData) {
  for (var r = ROWS - 1; r >= 0; r--) {
    if (boardData[r].every(val => val !== 0)) {
      boardData.splice(r, 1);
      boardData.unshift(new Array(COLS).fill(0));
      r++;
    }
  }
  return boardData;
}

function render(boardData, currentPiece, currentX, currentY, isPlaying) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tetris");
  if (!sheet) return;
  var displayGrid = boardData.map(row => row.slice());
  
  if (currentPiece && isPlaying) {
    for (var r = 0; r < currentPiece.length; r++) {
      for (var c = 0; c < currentPiece[r].length; c++) {
        if (currentPiece[r][c] !== 0) {
          if (currentY + r >= 0 && currentY + r < ROWS) {
            displayGrid[currentY + r][currentX + c] = "P"; 
          }
        }
      }
    }
  }

  var outputValues = displayGrid.map(row => row.map(val => val === 0 ? "" : val));
  sheet.getRange(1, 2, ROWS, COLS).setValues(outputValues);
  SpreadsheetApp.flush();
}

function spawnPiece(boardData) {
  var keys = Object.keys(SHAPES);
  var randomKey = keys[Math.floor(Math.random() * keys.length)];
  var p = SHAPES[randomKey];
  var x = Math.floor((COLS - p.length) / 2);
  var y = 0;
  if (checkCollision(x, y, p, boardData)) {
    return { p: null, x: 0, y: 0, over: true };
  }
  return { p: p, x: x, y: y, over: false };
}

function startTetris() {
  var boardData = [];
  for (var r = 0; r < ROWS; r++) {
    boardData.push(new Array(COLS).fill(0));
  }
  var spawn = spawnPiece(boardData);
  if (spawn.over) {
    Browser.msgBox("Game Over!");
    return;
  }
  
  var currentPiece = spawn.p;
  var currentX = spawn.x;
  var currentY = spawn.y;
  var isPlaying = true;

  saveGameState(boardData, currentPiece, currentX, currentY, isPlaying);
  var startTime = new Date().getTime();
  var lastTick = startTime;

  while (isPlaying) {
    Utilities.sleep(50);
    var currentTime = new Date().getTime();
    var action = popInputQueue();
    if (action) {
      if (action === "LEFT" && !checkCollision(currentX - 1, currentY, currentPiece, boardData)) currentX--;
      if (action === "RIGHT" && !checkCollision(currentX + 1, currentY, currentPiece, boardData)) currentX++;
      if (action === "DROP" && !checkCollision(currentX, currentY + 1, currentPiece, boardData)) currentY++;
      if (action === "ROTATE") {
        var rotated = rotateMatrix(currentPiece);
        if (!checkCollision(currentX, currentY, rotated, boardData)) currentPiece = rotated;
      }
      render(boardData, currentPiece, currentX, currentY, isPlaying);
      saveGameState(boardData, currentPiece, currentX, currentY, isPlaying);
    }
    if (currentTime - lastTick >= TICK_SPEED_MS) {
      if (!checkCollision(currentX, currentY + 1, currentPiece, boardData)) {
        currentY++;
      } else {
        for (var r = 0; r < currentPiece.length; r++) {
          for (var c = 0; c < currentPiece[r].length; c++) {
            if (currentPiece[r][c] !== 0 && (currentY + r) >= 0) {
              boardData[currentY + r][currentX + c] = 1;
            }
          }
        }
        boardData = clearLines(boardData);
        
        var nextSpawn = spawnPiece(boardData);
        if (nextSpawn.over) {
          isPlaying = false;
          render(boardData, null, 0, 0, false);
          saveGameState(boardData, null, 0, 0, false);
          Browser.msgBox("Game Over!");
          break;
        }
        currentPiece = nextSpawn.p;
        currentX = nextSpawn.x;
        currentY = nextSpawn.y;
      }
      render(boardData, currentPiece, currentX, currentY, isPlaying);
      saveGameState(boardData, currentPiece, currentX, currentY, isPlaying);
      lastTick = currentTime;
    }
  }
}
