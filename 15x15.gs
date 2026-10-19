const START_ROW = 3;
const TOTAL_ROWS = 15;
const NAME_COL = 23;  
const SCORE_COL = 24;
const ABC = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O"];

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function getCache() {
  return CacheService.getScriptCache();
}

function normalizeHex(color, defaultHex) {
  if (!color || typeof color !== "string") return defaultHex;
  let clean = color.trim();
  if (clean.startsWith("#")) clean = clean.substring(1);
  if (clean.length === 3) {
    clean = clean.split("").map(c => c + c).join("");
  }
  if (clean.length === 6 && /^[0-9A-Fa-f]{6}$/.test(clean)) {
    return "#" + clean.toLowerCase();
  }
  return clean.toLowerCase() || defaultHex;
}

function getHighScores() {
  const sheet = getSheet();
  const names = sheet.getRange(START_ROW, NAME_COL, TOTAL_ROWS, 1).getValues();
  const scores = sheet.getRange(START_ROW, SCORE_COL, TOTAL_ROWS, 1).getValues();
  let highScores = [];

  for (let i = 0; i < TOTAL_ROWS; i++) {
    highScores.push({
      name: String(names[i][0] || "Empty").trim(),
      hScore: Number(scores[i][0] || 0)
    });
  }

  highScores.sort((a, b) => b.hScore - a.hScore);
  return highScores;
}

function death(score, dead) {
  const sheet = getSheet();
  const ui = SpreadsheetApp.getUi();
  const highScores = getHighScores();

  let lowestCurrentScore = highScores[highScores.length - 1].hScore;

  if (score > lowestCurrentScore && score > 0) {
    const targetIndex = highScores.findIndex(cscore => score > cscore.hScore);
    const place = targetIndex + 1;
    let numplace = (place === 1) ? "st" : (place === 2) ? "nd" : (place === 3) ? "rd" : "th";
    const newName = sheet.getRange("S3").getValue() || "Player";

    highScores.splice(targetIndex, 0, { name: newName, hScore: score });
    highScores.pop();

    sheet.getRange(START_ROW, NAME_COL, TOTAL_ROWS, 1).setValues(highScores.map(i => [i.name]));
    sheet.getRange(START_ROW, SCORE_COL, TOTAL_ROWS, 1).setValues(highScores.map(i => [i.hScore]));

    ui.alert("New Highscore! You got " + place + numplace + " place with a score of " + score + "!");
  } else {
    ui.alert("Game Over! Your score: " + score);
  }
}

function respawnApple(snake, existingApples, walls) {
  let newApple;
  let collision;
  do {
    collision = false;
    newApple = {
      x: Math.floor(Math.random() * 15),
      y: Math.floor(Math.random() * 15)
    };

    for (let i = 0; i < snake.length; i++) {
      if (snake[i].x === newApple.x && snake[i].y === newApple.y) {
        collision = true;
        break;
      }
    }

    if (!collision) {
      for (let j = 0; j < existingApples.length; j++) {
        if (existingApples[j].x === newApple.x && existingApples[j].y === newApple.y) {
          collision = true;
          break;
        }
      }
    }

    if (!collision && walls) {
      for (let w = 0; w < walls.length; w++) {
        if (walls[w].x === newApple.x && walls[w].y === newApple.y) {
          collision = true;
          break;
        }
      }
    }
  } while (collision);

  return newApple;
}

function setupGame15() {
  const sheet = getSheet();
  const cache = getCache();

  sheet.getRange("A16:O16").setBackground("white");
  sheet.getRange("P1:P16").setBackground("white");
  sheet.getRange("P16").setValue("");

  const bgColor = normalizeHex(sheet.getRange("T27").getValue(), "#000000");
  const grid = Array.from({ length: 15 }, () => Array(15).fill(bgColor));
  sheet.getRange(1, 1, 15, 15).setBackgrounds(grid);

  const snake = [{x: 7, y: 7}];
  const apples = [];
  const walls = [];

  const useMapData = Number(sheet.getRange("T29").getValue()) === 1;

  if (useMapData) {
    const mapBackgrounds = sheet.getRange("A20:O34").getBackgrounds();
    for (let r = 0; r < 15; r++) {
      for (let c = 0; c < 15; c++) {
        let hex = normalizeHex(mapBackgrounds[r][c], "#ffffff");
        if (hex === "#ff0000" || hex === "#red") {
          if (!(c === snake[0].x && r === snake[0].y)) {
            apples.push({ x: c, y: r });
          }
        } else if (hex === "#000000" || hex === "#black") {
          if (!(c === snake[0].x && r === snake[0].y)) {
            walls.push({ x: c, y: r });
          }
        }
      }
    }
  } else {
    let rawVal = sheet.getRange("T22").getValue();
    let requestedCount = parseInt(rawVal, 10);
    if (isNaN(requestedCount) || requestedCount < 1) requestedCount = 1;
    if (requestedCount > 224) requestedCount = 224;

    for (let k = 0; k < requestedCount; k++) {
      apples.push(respawnApple(snake, apples, walls));
    }
  }

  let score = 0;

  cache.put("dir", "RIGHT");
  sheet.getRange("R2").setValue(score);

  runEngineLoop(sheet, snake, apples, walls, score);
}

function runEngineLoop(sheet, snake, apples, walls, score) {
  const cache = getCache();
  const maxTicks = 150;
  const dead = "N";

  let controls = sheet.getRange("Q6:S12").getValues();
  let upCell = controls[0][1];
  let leftCell = controls[3][0];
  let rightCell = controls[3][2];
  let downCell = controls[6][1];

  const appleColor = normalizeHex(sheet.getRange("T24").getValue(), "#FF0000");
  const bodyColor = normalizeHex(sheet.getRange("T25").getValue(), "#00FF00");
  const headColor = normalizeHex(sheet.getRange("T26").getValue(), "#008000");
  const bgColor = normalizeHex(sheet.getRange("T27").getValue(), "#000000");
  const wallColor = "#808080";
  const speed = Number(sheet.getRange("T21").getValue()) || 200;
  const disableWalls = Number(sheet.getRange("T28").getValue()) === 1;
  const reverseOnEat = Number(sheet.getRange("T31").getValue()) === 1;

  const baseGrid = Array.from({ length: 15 }, () => Array(15).fill(bgColor));

  for (let tick = 0; tick < maxTicks; tick++) {
    controls = sheet.getRange("Q6:S12").getValues();
    let curUp = controls[0][1];
    let curLeft = controls[3][0];
    let curRight = controls[3][2];
    let curDown = controls[6][1];

    let currentDir = cache.get("dir") || "RIGHT";

    if (curUp !== upCell && currentDir !== "DOWN") { upCell = curUp; cache.put("dir", "UP"); currentDir = "UP"; }
    if (curDown !== downCell && currentDir !== "UP") { downCell = curDown; cache.put("dir", "DOWN"); currentDir = "DOWN"; }
    if (curLeft !== leftCell && currentDir !== "RIGHT") { leftCell = curLeft; cache.put("dir", "LEFT"); currentDir = "LEFT"; }
    if (curRight !== rightCell && currentDir !== "LEFT") { rightCell = curRight; cache.put("dir", "RIGHT"); currentDir = "RIGHT"; }

    let head = {x: snake[0].x, y: snake[0].y};
    if (currentDir === "UP")    head.y--;
    if (currentDir === "DOWN")  head.y++;
    if (currentDir === "LEFT")  head.x--;
    if (currentDir === "RIGHT") head.x++;

    if (disableWalls) {
      if (head.x < 0) head.x = 14;
      if (head.x >= 15) head.x = 0;
      if (head.y < 0) head.y = 14;
      if (head.y >= 15) head.y = 0;
    } else {
      head.lx = (head.x < 0) ? "A" : (head.x >= 15) ? "O" : ABC[head.x];
      if (head.x < 0 || head.x >= 15 || head.y < 0 || head.y >= 15) {
        sheet.getRange(head.lx + "16").setBackground("green");
        death(score, dead);
        return;
      }
    }

    for (let w = 0; w < walls.length; w++) {
      if (walls[w].x === head.x && walls[w].y === head.y) {
        death(score, dead);
        return;
      }
    }

    for (let i = 0; i < snake.length; i++) {
      if (snake[i].x === head.x && snake[i].y === head.y) {
        death(score, dead);
        return;
      }
    }

    snake.unshift(head);

    let eatenIndex = apples.findIndex(a => a.x === head.x && a.y === head.y);
    if (eatenIndex !== -1) {
      score += 1;
      sheet.getRange("R2").setValue(score);
      apples[eatenIndex] = respawnApple(snake, apples, walls);

      if (reverseOnEat && snake.length > 1) {
        snake.reverse();
        let newHead = snake[0];
        let newNeck = snake[1];

        let dx = newHead.x - newNeck.x;
        let dy = newHead.y - newNeck.y;

        if (disableWalls) {
          if (dx === -14) dx = 1;
          else if (dx === 14) dx = -1;
          if (dy === -14) dy = 1;
          else if (dy === 14) dy = -1;
        }

        let newDir = currentDir;
        if (dx === 1) newDir = "RIGHT";
        else if (dx === -1) newDir = "LEFT";
        else if (dy === 1) newDir = "DOWN";
        else if (dy === -1) newDir = "UP";

        cache.put("dir", newDir);
      }
    } else {
      snake.pop();
    }

    let frameColors = baseGrid.map(row => row.slice());

    for (let w = 0; w < walls.length; w++) {
      frameColors[walls[w].y][walls[w].x] = wallColor;
    }

    for (let a = 0; a < apples.length; a++) {
      frameColors[apples[a].y][apples[a].x] = appleColor;
    }

    for (let s = 0; s < snake.length; s++) {
      if (s === 0) {
        frameColors[snake[s].y][snake[s].x] = headColor;
      } else {
        frameColors[snake[s].y][snake[s].x] = bodyColor;
      }
    }

    sheet.getRange(1, 1, 15, 15).setBackgrounds(frameColors);
    SpreadsheetApp.flush();
    Utilities.sleep(speed);
  }
}
