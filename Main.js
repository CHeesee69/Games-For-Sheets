
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var CACHE = CacheService.getScriptCache();

const START_ROW = 3;
const TOTAL_ROWS = 15;
const NAME_COL = 23;  
const SCORE_COL = 24; 

function getHighScores() {
  const namesColumn = sheet.getRange(START_ROW, NAME_COL, TOTAL_ROWS, 1).getValues();
  const scoresColumn = sheet.getRange(START_ROW, SCORE_COL, TOTAL_ROWS, 1).getValues();

  let highScores = [];
  for (let i = 0; i < TOTAL_ROWS; i++) {
    highScores.push({
      name: String(namesColumn[i][0] || "Empty").trim(),
      hScore: Number(scoresColumn[i][0] || 0)
    });
  }

  highScores.sort((a, b) => b.hScore - a.hScore);
  return highScores;
}

function death(score, dead) {
  dead = "Y";
  const ui = SpreadsheetApp.getUi();
  const highScores = getHighScores();
  let lowestCurrentScore = highScores[highScores.length - 1].hScore;
  
  if (score > lowestCurrentScore && score > 0) {
    const targetIndex = highScores.findIndex(cscore => score > cscore.hScore);
    const place = targetIndex + 1;
    
    let numplace = "th";
    if (place === 1) numplace = "st";
    else if (place === 2) numplace = "nd";
    else if (place === 3) numplace = "rd";
    
    const newName = sheet.getRange("S3").getValue() || "Player";

    highScores.splice(targetIndex, 0, { name: newName, hScore: score });
    highScores.pop();
    
    const updatedNames = highScores.map(item => [item.name]);
    const updatedScores = highScores.map(item => [item.hScore]);
    
    sheet.getRange(START_ROW, NAME_COL, TOTAL_ROWS, 1).setValues(updatedNames);
    sheet.getRange(START_ROW, SCORE_COL, TOTAL_ROWS, 1).setValues(updatedScores);
    
    ui.alert("New Highscore! You got " + place + numplace + " place with a score of " + score + "!");
  } else {
    ui.alert("Game Over! Your score: " + score);
  }
}

function setupGame() {
  var dead = "N";
  
  sheet.getRange("A16:O16").setBackground("white");
  sheet.getRange("P1:P16").setBackground("white");
  sheet.getRange("P16").setValue("");
  
  var blackGrid = [];
  for (var i = 0; i < 15; i++) {
    var row = [];
    for (var j = 0; j < 15; j++) { row.push("#000000"); }
    blackGrid.push(row);
  }
  sheet.getRange(1, 1, 15, 15).setBackgrounds(blackGrid);

  var snake = [{x: 7, y: 7}];
  var apple = {x: 3, y: 3};
  var score = 0;
  CACHE.put("dir", "RIGHT");
  sheet.getRange("R2").setValue(score);
  
  runEngineLoop(sheet, snake, apple, score);
}

function runEngineLoop(sheet, snake, apple, score) {
  var maxTicks = 150;
  var dead = "N";
  
  var upCell = sheet.getRange("R6").getValue();
  var downCell = sheet.getRange("R12").getValue();
  var leftCell = sheet.getRange("Q9").getValue();
  var rightCell = sheet.getRange("S9").getValue();
  
  for (var tick = 0; tick < maxTicks; tick++) {

    var curUp = sheet.getRange("R6").getValue();
    var curDown = sheet.getRange("R12").getValue();
    var curLeft = sheet.getRange("Q9").getValue();
    var curRight = sheet.getRange("S9").getValue();
    var currentDir = CACHE.get("dir") || "RIGHT";
    if (curUp !== upCell && currentDir !== "DOWN") { upCell = curUp; CACHE.put("dir", "UP"); }
    if (curDown !== downCell && currentDir !== "UP") { downCell = curDown; CACHE.put("dir", "DOWN"); }
    if (curLeft !== leftCell && currentDir !== "RIGHT") { leftCell = curLeft; CACHE.put("dir", "LEFT"); }
    if (curRight !== rightCell && currentDir !== "LEFT") { rightCell = curRight; CACHE.put("dir", "RIGHT"); }
    var head = {x: snake[0].x, y: snake[0].y};
    
    if (currentDir === "UP")    head.y--;
    if (currentDir === "DOWN")  head.y++;
    if (currentDir === "LEFT")  head.x--;
    if (currentDir === "RIGHT") head.x++;
    
    const abc = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O"];
    
    if (head.x < 0) head.lx = "A";
    else if (head.x >= 15) head.lx = "O";
    else head.lx = abc[head.x];

    if (head.x < 0 || head.x >= 15) {
      sheet.getRange(head.lx + "16").setBackground("green");
      sheet.getRange("P" + (head.y + 1)).setBackground("green"); 
      sheet.getRange("P16").setValue(head.lx + (head.y + 1));
      death(score, dead); 
      return;
    }

    if (head.y < 0 || head.y >= 15) {
      var topBottom = (head.y >= 15) ? 0 : 2;
      sheet.getRange(head.lx + "16").setBackground("green");
      sheet.getRange("P" + (head.y + topBottom)).setBackground("green"); 
      sheet.getRange("P16").setValue(head.lx + (head.y + topBottom));
      death(score, dead); 
      return;
    }

    for (var i = 0; i < snake.length; i++) {
      if (snake[i].x === head.x && snake[i].y === head.y) {
        sheet.getRange(`${head.lx}16`).setBackground("blue");
        sheet.getRange("P" + (head.y + 1)).setBackground("blue"); 
        sheet.getRange("P16").setValue(head.lx + (head.y + 1));
        death(score, dead); 
        return;
      }
    }
    
    snake.unshift(head);
    
    if (head.x === apple.x && head.y === apple.y) {
      score += 1;
      sheet.getRange("R2").setValue(score);
   
      apple.x = Math.floor(Math.random() * 15);
      apple.y = Math.floor(Math.random() * 15);
    } else {
      snake.pop(); 
    }
    var frameColors = [];
    for (var r = 0; r < 15; r++) {
      var rowColors = [];
      for (var c = 0; c < 15; c++) { rowColors.push("#000000"); }
      frameColors.push(rowColors);
    }
    
    frameColors[apple.y][apple.x] = "#FF0000"; 
    snake.forEach(function(segment) {
      frameColors[segment.y][segment.x] = "#00FF00"; 
    });
    
    sheet.getRange(1, 1, 15, 15).setBackgrounds(frameColors);
    SpreadsheetApp.flush();
    Utilities.sleep(350); 
  }
}
