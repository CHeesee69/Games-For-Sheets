const Snake10x10 = {
  START_ROW: 3,
  TOTAL_ROWS: 10,
  NAME_COL: 23,
  SCORE_COL: 24,
  ABC: ["A","B","C","D","E","F","G","H","I","J"],

  getSheet: function() {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  },

  getCache: function() {
    return CacheService.getScriptCache();
  },

  normalizeHex: function(color, defaultHex) {
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
  },

  getHighScores: function() {
    const sheet = this.getSheet();
    const names = sheet.getRange(this.START_ROW, this.NAME_COL, this.TOTAL_ROWS, 1).getValues();
    const scores = sheet.getRange(this.START_ROW, this.SCORE_COL, this.TOTAL_ROWS, 1).getValues();
    let highScores = [];

    for (let i = 0; i < this.TOTAL_ROWS; i++) {
      highScores.push({
        name: String(names[i][0] || "Empty").trim(),
        hScore: Number(scores[i][0] || 0)
      });
    }

    highScores.sort((a, b) => b.hScore - a.hScore);
    return highScores;
  },

  death: function(score, dead) {
    const sheet = this.getSheet();
    const ui = SpreadsheetApp.getUi();
    const highScores = this.getHighScores();

    let lowestCurrentScore = highScores[highScores.length - 1].hScore;

    if (score > lowestCurrentScore && score > 0) {
      const targetIndex = highScores.findIndex(cscore => score > cscore.hScore);
      const place = targetIndex + 1;
      let numplace = (place === 1) ? "st" : (place === 2) ? "nd" : (place === 3) ? "rd" : "th";
      const newName = sheet.getRange("S3").getValue() || "Player";

      highScores.splice(targetIndex, 0, { name: newName, hScore: score });
      highScores.pop();

      sheet.getRange(this.START_ROW, this.NAME_COL, this.TOTAL_ROWS, 1).setValues(highScores.map(i => [i.name]));
      sheet.getRange(this.START_ROW, this.SCORE_COL, this.TOTAL_ROWS, 1).setValues(highScores.map(i => [i.hScore]));

      ui.alert("New Highscore! You got " + place + numplace + " place with a score of " + score + "!");
    } else {
      ui.alert("Game Over! Your score: " + score);
    }
  },

  respawnApple: function(snake, existingApples, walls) {
    let newApple;
    let collision;
    do {
      collision = false;
      newApple = {
        x: Math.floor(Math.random() * 10),
        y: Math.floor(Math.random() * 10)
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
  },

  setupGame: function() {
    const sheet = this.getSheet();
    const cache = this.getCache();

    sheet.getRange("A11:J11").setBackground("white");
    sheet.getRange("K1:K11").setBackground("white");
    sheet.getRange("K11").setValue("");

    const bgColor = this.normalizeHex(sheet.getRange("T27").getValue(), "#000000");
    const grid = Array.from({ length: 10 }, () => Array(10).fill(bgColor));
    sheet.getRange(1, 1, 10, 10).setBackgrounds(grid);

    const snake = [{x: 2, y: 4}];
    const apples = [];
    const walls = [];

    const useMapData = Number(sheet.getRange("T29").getValue()) === 1;

    if (useMapData) {
      const mapBackgrounds = sheet.getRange("A20:J29").getBackgrounds();
      for (let r = 0; r < 10; r++) {
        for (let c = 0; c < 10; c++) {
          let hex = this.normalizeHex(mapBackgrounds[r][c], "#ffffff");
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
      if (requestedCount > 99) requestedCount = 99;

      for (let k = 0; k < requestedCount; k++) {
        apples.push(this.respawnApple(snake, apples, walls));
      }
    }

    let score = 0;

    cache.put("dir_10x10", "RIGHT");
    sheet.getRange("R2").setValue(score);

    this.runEngineLoop(sheet, snake, apples, walls, score);
  },

  runEngineLoop: function(sheet, snake, apples, walls, score) {
    const cache = this.getCache();
    const maxTicks = 150;
    const dead = "N";

    let controls = sheet.getRange("Q6:S12").getValues();
    let upCell = controls[0][1];
    let leftCell = controls[3][0];
    let rightCell = controls[3][2];
    let downCell = controls[6][1];

    const appleColor = this.normalizeHex(sheet.getRange("T24").getValue(), "#FF0000");
    const bodyColor = this.normalizeHex(sheet.getRange("T25").getValue(), "#00FF00");
    const headColor = this.normalizeHex(sheet.getRange("T26").getValue(), "#008000");
    const bgColor = this.normalizeHex(sheet.getRange("T27").getValue(), "#000000");
    const wallColor = "#808080";
    const speed = Number(sheet.getRange("T21").getValue()) || 200;
    const disableWalls = Number(sheet.getRange("T28").getValue()) === 1;
    const reverseOnEat = Number(sheet.getRange("T31").getValue()) === 1;

    const baseGrid = Array.from({ length: 10 }, () => Array(10).fill(bgColor));

    for (let tick = 0; tick < maxTicks; tick++) {
      controls = sheet.getRange("Q6:S12").getValues();
      let curUp = controls[0][1];
      let curLeft = controls[3][0];
      let curRight = controls[3][2];
      let curDown = controls[6][1];

      let currentDir = cache.get("dir_10x10") || "RIGHT";

      if (curUp !== upCell && currentDir !== "DOWN") { upCell = curUp; cache.put("dir_10x10", "UP"); currentDir = "UP"; }
      if (curDown !== downCell && currentDir !== "UP") { downCell = curDown; cache.put("dir_10x10", "DOWN"); currentDir = "DOWN"; }
      if (curLeft !== leftCell && currentDir !== "RIGHT") { leftCell = curLeft; cache.put("dir_10x10", "LEFT"); currentDir = "LEFT"; }
      if (curRight !== rightCell && currentDir !== "LEFT") { rightCell = curRight; cache.put("dir_10x10", "RIGHT"); currentDir = "RIGHT"; }

      let head = {x: snake[0].x, y: snake[0].y};
      if (currentDir === "UP")    head.y--;
      if (currentDir === "DOWN")  head.y++;
      if (currentDir === "LEFT")  head.x--;
      if (currentDir === "RIGHT") head.x++;

      if (disableWalls) {
        if (head.x < 0) head.x = 9;
        if (head.x >= 10) head.x = 0;
        if (head.y < 0) head.y = 9;
        if (head.y >= 10) head.y = 0;
      } else {
        head.lx = (head.x < 0) ? "A" : (head.x >= 10) ? "J" : this.ABC[head.x];
        if (head.x < 0 || head.x >= 10 || head.y < 0 || head.y >= 10) {
          sheet.getRange(head.lx + "11").setBackground("green");
          this.death(score, dead);
          return;
        }
      }

      for (let w = 0; w < walls.length; w++) {
        if (walls[w].x === head.x && walls[w].y === head.y) {
          this.death(score, dead);
          return;
        }
      }

      for (let i = 0; i < snake.length; i++) {
        if (snake[i].x === head.x && snake[i].y === head.y) {
          this.death(score, dead);
          return;
        }
      }

      snake.unshift(head);

      let eatenIndex = apples.findIndex(a => a.x === head.x && a.y === head.y);
      if (eatenIndex !== -1) {
        score += 1;
        sheet.getRange("R2").setValue(score);
        apples[eatenIndex] = this.respawnApple(snake, apples, walls);

        if (reverseOnEat && snake.length > 1) {
          snake.reverse();
          let newHead = snake[0];
          let newNeck = snake[1];

          let dx = newHead.x - newNeck.x;
          let dy = newHead.y - newNeck.y;

          if (disableWalls) {
            if (dx === -9) dx = 1;
            else if (dx === 9) dx = -1;
            if (dy === -9) dy = 1;
            else if (dy === 9) dy = -1;
          }

          let newDir = currentDir;
          if (dx === 1) newDir = "RIGHT";
          else if (dx === -1) newDir = "LEFT";
          else if (dy === 1) newDir = "DOWN";
          else if (dy === -1) newDir = "UP";

          cache.put("dir_10x10", newDir);
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

      sheet.getRange(1, 1, 10, 10).setBackgrounds(frameColors);
      SpreadsheetApp.flush();
      Utilities.sleep(speed);
    }
  }
};

function runSnake10x10() {
  Snake10x10.setupGame();
}