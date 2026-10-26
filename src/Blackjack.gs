const cards = [
  { card: "Ac", sheetloc: "B38"},  { card: "Ad", sheetloc: "L38"},  { card: "Ah", sheetloc: "V38"},  { card: "As", sheetloc: "AF38"},
  { card: "2c", sheetloc: "B54"},  { card: "2d", sheetloc: "L54"},  { card: "2h", sheetloc: "V54"},  { card: "2s", sheetloc: "AF54"},
  { card: "3c", sheetloc: "B70"},  { card: "3d", sheetloc: "L70"},  { card: "3h", sheetloc: "V70"},  { card: "3s", sheetloc: "AF70"},
  { card: "4c", sheetloc: "B86"},  { card: "4d", sheetloc: "L86"},  { card: "4h", sheetloc: "V86"},  { card: "4s", sheetloc: "AF86"},
  { card: "5c", sheetloc: "B102"}, { card: "5d", sheetloc: "L102"}, { card: "5h", sheetloc: "V102"}, { card: "5s", sheetloc: "AF102"},
  { card: "6c", sheetloc: "B118"}, { card: "6d", sheetloc: "L118"}, { card: "6h", sheetloc: "V118"}, { card: "6s", sheetloc: "AF118"},
  { card: "7c", sheetloc: "B134"}, { card: "7d", sheetloc: "L134"}, { card: "7h", sheetloc: "V134"}, { card: "7s", sheetloc: "AF134"},
  { card: "8c", sheetloc: "B150"}, { card: "8d", sheetloc: "L150"}, { card: "8h", sheetloc: "V150"}, { card: "8s", sheetloc: "AF150"},
  { card: "9c", sheetloc: "B166"}, { card: "9d", sheetloc: "L166"}, { card: "9h", sheetloc: "V166"}, { card: "9s", sheetloc: "AF166"},
  { card: "10c", sheetloc: "B182"},{ card: "10d", sheetloc: "L182"},{ card: "10h", sheetloc: "V182"},{ card: "10s", sheetloc: "AF182"},
  { card: "Jc", sheetloc: "B198"}, { card: "Jd", sheetloc: "L198"}, { card: "Jh", sheetloc: "V198"}, { card: "Js", sheetloc: "AF198"},
  { card: "Qc", sheetloc: "B214"}, { card: "Qd", sheetloc: "L214"}, { card: "Qh", sheetloc: "V214"}, { card: "Qs", sheetloc: "AF214"},
  { card: "Kc", sheetloc: "B230"}, { card: "Kd", sheetloc: "L230"}, { card: "Kh", sheetloc: "V230"}, { card: "Ks", sheetloc: "AF230"}
];
const PLACARLOC = ["B3:J17", "L3:T17", "V3:AD17", "AF3:AN17", "AP3:AX17"];
const DELCARLOC = ["B20:J34", "L20:T34", "V20:AD34", "AF20:AN34", "AP20:AX34"];
function startBlackjack() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Blackjack");
  const props = PropertiesService.getScriptProperties();
  let color = sheet.getRange("BA18").getValue();
  PLACARLOC.concat(DELCARLOC).forEach(range => sheet.getRange(range).clearContent());
  let deck = cards.map(c => c.card);
  for (let i = deck.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [deck[i], deck[j]] = [deck[j], deck[i]];
  }
  const pCard1 = deck.shift();
  const dCard1 = deck.shift();
  const pCard2 = deck.shift();
  const dCard2 = deck.shift();
  const playerHand = [pCard1, pCard2];
  cretCard(pCard1, PLACARLOC[0]);
  cretCard(pCard2, PLACARLOC[1]);
  cretCard(dCard1, DELCARLOC[0]);
  sheet.getRange(color).copyTo(sheet.getRange(DELCARLOC[1]));
  sheet.getRange("AJ1").setValue(add(playerHand));
  sheet.getRange("AJ18").setValue(add([dCard1]));
  props.setProperty("DECK", JSON.stringify(deck));
  props.setProperty("PLARHND", JSON.stringify(playerHand));
  props.setProperty("DELHND", JSON.stringify([dCard1, dCard2]));
  props.setProperty("GAMEAC", "true");
}
function hit() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Blackjack");
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("GAMEAC") !== "true") return;
  let deck = JSON.parse(props.getProperty("DECK"));
  let playerHand = JSON.parse(props.getProperty("PLARHND"));

  if (playerHand.length >= PLACARLOC.length) {
    return;
  }
  const newCard = deck.shift();
  playerHand.push(newCard);
  const targetRange = PLACARLOC[playerHand.length - 1];
  cretCard(newCard, targetRange);
  const currentScore = add(playerHand);
  sheet.getRange("AJ1").setValue(currentScore);
  props.setProperty("DECK", JSON.stringify(deck));
  props.setProperty("PLARHND", JSON.stringify(playerHand));

  if (currentScore > 21) {
    props.setProperty("GAMEAC", "false");
    revl();
    let dealerHand = JSON.parse(props.getProperty("DELHND"));
    sheet.getRange("AJ18").setValue(add(dealerHand));
    SpreadsheetApp.flush();
    Utilities.sleep(500);
    Browser.msgBox("Bust!");
  }
}

function stand() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Blackjack");
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("GAMEAC") !== "true") return;
  props.setProperty("GAMEAC", "false");
  revl();
  let deck = JSON.parse(props.getProperty("DECK"));
  let dealerHand = JSON.parse(props.getProperty("DELHND"));
  let playerHand = JSON.parse(props.getProperty("PLARHND"));
  while (add(dealerHand) < 17 && dealerHand.length < DELCARLOC.length) {
    const newCard = deck.shift();
    dealerHand.push(newCard);
    cretCard(newCard, DELCARLOC[dealerHand.length - 1]);
  }
  const dealerScore = add(dealerHand);
  sheet.getRange("AJ18").setValue(dealerScore);
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  const playerScore = add(playerHand);
  if (dealerScore > 21 || playerScore > dealerScore) {
    Browser.msgBox(`You Won!`);
  } else if (dealerScore > playerScore) {
    Browser.msgBox(`Dealer Wins!`);
  } else {
    Browser.msgBox(`Push! ${playerScore}`);
  }
}

function cretCard(cardCode, targetRange) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Blackjack");
  const cardObj = cards.find(c => c.card === cardCode);
  if (cardObj) {
    sheet.getRange(cardObj.sheetloc).copyTo(sheet.getRange(targetRange));
  }
}

function revl() {
  const props = PropertiesService.getScriptProperties();
  const dealerHand = JSON.parse(props.getProperty("DELHND"));
  cretCard(dealerHand[1], DELCARLOC[1]);
}

function add(hand) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Blackjack");
  let score = 0;
  let aces = 0;

  for (let cardCode of hand) {
    let rank = cardCode.slice(0, -1);
    if (rank === "A") {
      score += 11;
      aces++;
    } else if (["J", "Q", "K"].includes(rank)) {
      score += 10;
    } else {
      score += parseInt(rank, 10);
    }
  }

  while (score > 21 && aces > 0) {
    score -= 10;
    aces--;
  }

  return score;
}
