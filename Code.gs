function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Reload Item Prices', 'reloadPrices')
      .addToUi();
}

const API_USERNAME = "Refraction"

const itemsColumnAvgPrices = letterToNumber("K")

const itemsColumnItemID = letterToNumber("B")
const itemsColumnItemBoughtQTY = letterToNumber("C")
const itemsColumnItemSoldQTY = letterToNumber("G")

const itemsStartRow = 2

function reloadPrices() {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let itemsS = ss.getSheetByName("items")
  if (!itemsS) {
    throw "Error: couldn't find 'items' sheet"
  }
  let priceS = ss.getSheetByName("price_history")
  if (!priceS) {
    throw "Error: couldn't find 'price_history' sheet"
  }
  listItems()
}

/**
 * @returns {Object} - a map with the key being items and the value being the amount of said items
 */
function listItems() {
  let itemsS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("items")
  if (!itemsS) {
    throw "Error: couldn't find 'items' sheet"
  }
  let maxRow = itemsS.getDataRange().getNumRows()
  if (maxRow ===  1) {
    throw "Error: couldn't find any items"
  }
  let range = itemsS.getRange(itemsStartRow,itemsColumnItemID,maxRow,itemsColumnItemSoldQTY-itemsColumnItemID+1)
  let values = range.getValues()
  let r = {}
  for (let i = 0;i < maxRow;i++) {
    let item_id = values[i][0]
    if (!item_id) continue

    let boughtQTY = values[i][itemsColumnItemBoughtQTY-itemsColumnItemID]
    if (!boughtQTY) continue
    boughtQTY = parseInt(boughtQTY)
    if (!boughtQTY) continue

    let soldQTY = values[i][itemsColumnItemSoldQTY-itemsColumnItemID]
    if (soldQTY === undefined || soldQTY === "") soldQTY = 0
    soldQTY = parseInt(soldQTY)
    if (isNaN(soldQTY)) continue

    r[item_id] = {
      row: itemsStartRow+i,
      boughtQTY,
      soldQTY,
    }
  }
  throw JSON.stringify(r)
}

/**
 * API Stuff
 */
const SOOPY_API_URL = "https://soopy.dev/api/soopyv2/botcommand/"
const COFL_API_URL = "https://sky.coflnet.com/api/item/price/"

function getLowestBin(item_id) {
  const url = SOOPY_API_URL + `?m=lowestbin${encodeURIComponent(" ")}${encodeURIComponent(item_id)}&u=${encodeURIComponent(API_USERNAME)}`
  throw url
}


/**
 * Utils
 */
/**
 * @param {string} letter
 */
function letterToNumber(letter) {
  letter = letter.toLowerCase()
  let c = 0
  for(let i = 0;i < letter.length;i++) {
    let code = letter.charCodeAt(i);
    if (code<97 || code > 122) continue
    c += (code-97)
  }
  return c + 1 // columns start at 1
}
