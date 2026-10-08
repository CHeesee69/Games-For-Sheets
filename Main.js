 const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getRange("A16");
  cell.setBackground("white"); 
var CACHE = CacheService.getScriptCache();

function setupGame() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  

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
  
  for (var tick = 0; tick < maxTicks; tick++) {

    var currentDir = CACHE.get("dir") || "RIGHT";
    var head = {x: snake[0].x, y: snake[0].y};
    
    if (currentDir === "UP")    head.y--;
    if (currentDir === "DOWN")  head.y++;
    if (currentDir === "LEFT")  head.x--;
    if (currentDir === "RIGHT") head.x++;
    
        if (head.x < 0 || head.x >= 15 || head.y < 0 || head.y >= 15) {
          const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
          const cell = sheet.getRange("A16");
          cell.setBackground("green"); ;
      return;
    }

    for (var i = 0; i < snake.length; i++) {
      if (snake[i].x === head.x && snake[i].y === head.y) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const cell = sheet.getRange("A16");
        cell.setBackground("blue"); 
        return;
      }
    }
    
    snake.unshift(head);
    
    if (head.x === apple.x && head.y === apple.y) {
      score += 10;
      sheet.getRange("R2").setValue(score);
   
      apple.x = Math.floor(Math.random() * 15);
      apple.y = Math.floor(Math.random() * 15);
    } else {
      snake.pop(); // Remove tail
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

// Control assignments for your drawing shapes
function pressUp()    { CACHE.put("dir", "UP"); }
function pressDown()  { CACHE.put("dir", "DOWN"); }
function pressLeft()  { CACHE.put("dir", "LEFT"); }
function pressRight() { CACHE.put("dir", "RIGHT"); }
