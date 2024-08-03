function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Reload Item Prices', 'reloadPrices')
      .addToUi();
}

const API_USERNAME = "Refraction"

const itemsColumnAvgPrices = columnToNumber("K")
const itemsColumnSellVolume = columnToNumber("L")

const itemsColumnItemID = columnToNumber("B")
const itemsColumnItemBoughtQTY = columnToNumber("C")
const itemsColumnItemSoldQTY = columnToNumber("G")

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

  let items = listItems()
  for(let item_id in items) {
    let item = items[item_id]
    let resCell = itemsS.getRange(item.row,itemsColumnAvgPrices)
    try {
      let price = getAHInfo(item.runeID ? `UNIQUE_RUNE_${item.runeID}` : item_id)
      item.avg_price = price.avg
      item.volume = price.volume
      item.lbin = getLowestBin(item.runeID ?? item_id)

      resCell.setValue(Math.floor(price.avg/1000000))
      itemsS.getRange(item.row,itemsColumnSellVolume).setValue(price.volume)
    } catch (err) {
      resCell.setValue(err.toString())
    }
  }
  console.log(items)
}

/**
 * @returns {Object} - a map with the key being items and the value being data about the item
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
    /**
     * @type {string}
     */
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
      id: item_id,
      row: itemsStartRow+i,
      boughtQTY,
      soldQTY,
      runeID: item_id.startsWith("UNIQUE_RUNE.") ? item_id.split(".")[1] : undefined
    }
  }
  return r
}

/**
 * API Stuff
 */
const SOOPY_API_URL = "https://soopy.dev/api/soopyv2/botcommand/"
const COFL_API_URL = "https://sky.coflnet.com/api/item/price/"

function getLowestBin(item_id) {
  const url = SOOPY_API_URL + `?m=lowestbin${encodeURIComponent(" ")}${encodeURIComponent(item_id)}&u=${encodeURIComponent(API_USERNAME)}`
  let response = UrlFetchApp.fetch(url);
  let text = response.getContentText()
  if (!text.startsWith("Cheapest bin for")) throw ("Response is invalid: " + text)
  text = text.split(" ")
  text = text[text.length-1]
  text = text.substring(0,text.length-1)
  text = text.split(",").join("")
  text = parseInt(text)
  if (isNaN(text)) throw "Couldn't parse lowest bin number"
  return text
}

function getAHInfo(item_id) {
  const url = COFL_API_URL + `${encodeURIComponent(item_id)}/history/day`
  let response = UrlFetchApp.fetch(url);
  let r = JSON.parse(response.getContentText());
  if (!r) throw "Couldn't parse average price"
  if (r.slug) throw `${item_id}: ${r.message}`
  if (!Array.isArray(r)) throw "COFL reply isn't an array"
  let avg = 0
  let volume = 0
  for(let price_point of r) {
    avg += price_point.avg;
    volume += price_point.volume
  }
  return {avg: Math.floor(avg/r.length),volume}
}


/**
 * Utils
 */
/**
 * @param {string} column
 */
function columnToNumber(column) {
  column = column.toLowerCase()
  let index = 0;

  for (let i = 0; i < column.length; i++) {
    let code = column.charCodeAt(i);
    if (code<97 || code > 122) continue
    index = index * 26 + (code-97) + 1;
  }

  return index
}
