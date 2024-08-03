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

const itemsColumnBuyPrice = columnToNumber("D")
const itemsColumnSellPrice = columnToNumber("H")

const itemsStartRow = 2

const pricesColumnFirstData = columnToNumber("G")
const pricesOneItemWidth = 4 // col 1: item qty i still have, col 2: item lbin ; col3: item avg; col4: volume

function reloadPrices() {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let itemsS = ss.getSheetByName("items")
  if (!itemsS) {
    throw "Error: couldn't find 'items' sheet"
  }
  let items = listItems()
  // Fill in column in item prices
  for(let item_id in items) {
    let item = items[item_id]
    let resCell = itemsS.getRange(item.row,itemsColumnAvgPrices)
    try {
      let price = getAHInfo(item.runeID ? `UNIQUE_RUNE_${item.runeID}` : item_id)
      item.avg_price = price.avg
      item.volume = price.volume
      item.lbin = getLowestBin(item.runeID ?? item_id)

      resCell.setValue(Math.floor(price.avg/1e6))
      itemsS.getRange(item.row,itemsColumnSellVolume).setValue(price.volume)
    } catch (err) {
      item.failed_getting_cost = true
      resCell.setValue(err.toString())
    }
  }
  
  let priceS = ss.getSheetByName("price_history")
  if (!priceS) {
    throw "Error: couldn't find 'price_history' sheet"
  }
  let maxColumns = priceS.getDataRange().getNumColumns()
  let firstRow = priceS.getRange(1,1,1,maxColumns).getValues()[0]
  while(firstRow.length<pricesColumnFirstData-1) firstRow.push("")

  var new_row = new Array(pricesColumnFirstData-1)

  for(let i=pricesColumnFirstData;i<maxColumns;i+=pricesOneItemWidth) {
    let item_id = firstRow[i].split(" ")[1]
    if (item_id in items) {
      items[item_id].price_history_col = i
    } else {
      new_row[i-1] = 0
    }
  }
  
  let totalValueAverage = 0
  let totalValueLBIN = 0
  let totalProfitAverage = 0
  let totalPotentialProfit = 0


  for(let item_id in items) {
    let item = items[item_id]
    let i = item.price_history_col
    if (!i) {
      i = new_row.length
      for(let j = 0;j < pricesOneItemWidth;j++) new_row.push(0)
      firstRow.push(`| ${item_id} QTY |`)
      firstRow.push(`| ${item_id} LBIN |`)
      firstRow.push(`| ${item_id} AVG |`)
      firstRow.push(`| ${item_id} VOL |`)
    } else {
      i -= 1
    }
    let currentQTY = (item.boughtQTY - item.soldQTY)
    
    new_row[i] = currentQTY
    if (item.failed_getting_cost) continue
    new_row[i+1] = (Math.floor(item.lbin/1e6))
    new_row[i+2] = (Math.floor(item.avg_price/1e6))
    new_row[i+3] = (item.volume)

    

    totalValueAverage += item.avg_price * currentQTY
    totalValueLBIN += item.lbin * currentQTY
    totalProfitAverage += (item.avg_price * currentQTY) - (currentQTY*item.buyPrice) + ((item.sellPrice - item.buyPrice)*item.soldQTY)
    totalPotentialProfit += (item.avg_price * currentQTY) - (currentQTY*item.buyPrice)
  }

  if (new_row[new_row.length-1] === 0) new_row.pop()

  new_row[0] = new Date()
  new_row[1] = totalValueAverage
  new_row[2] = totalValueLBIN
  new_row[3] = totalProfitAverage
  new_row[4] = totalPotentialProfit

  if (firstRow.length !== maxColumns) {
    priceS.getRange(1,1,1,firstRow.length).setValues([firstRow])
    for(let i = Math.max(pricesColumnFirstData-1,maxColumns);i<firstRow.length;i+=pricesOneItemWidth) {
      let c = numberToColumn(i)
      priceS
        .getRange(`${c}:${c}`)
        .setBorder(null,true,null,false,null,null,"black", SpreadsheetApp.BorderStyle.SOLID)
    } 
  }
  priceS.appendRow(new_row)

  priceS.autoResizeColumns(pricesColumnFirstData,firstRow.length-pricesColumnFirstData)
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
  let range = itemsS.getRange(itemsStartRow,itemsColumnItemID,maxRow,itemsColumnSellPrice-itemsColumnItemID+1)
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

    let buyPrice = values[i][itemsColumnBuyPrice-itemsColumnItemID]
    if (buyPrice === undefined || buyPrice === "") continue
    buyPrice = parseInt(buyPrice)
    if (isNaN(buyPrice)) continue
    buyPrice = buyPrice * 1e6

    let sellPrice = values[i][itemsColumnSellPrice-itemsColumnItemID]
    if (sellPrice === undefined || sellPrice === "") sellPrice = 0
    sellPrice = parseInt(sellPrice)
    if (isNaN(sellPrice)) continue
    sellPrice = sellPrice * 1e6

    r[item_id] = {
      id: item_id,
      row: itemsStartRow+i,
      boughtQTY,
      soldQTY,
      buyPrice,
      sellPrice,
      runeID: item_id.startsWith("UNIQUE_RUNE.") ? item_id.split(".")[1] : undefined,
      price_history_col: null
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
 * Scripts
 */
function fetchPricesClick() {
  
  itemsS.activate()

  reloadPrices()
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

function numberToColumn(index) {
  const base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let column = '';

  while (index >= 0) {
    column = base[index % 26] + column;
    index = Math.floor(index / 26) - 1;
  }

  return column;
}

/**
 * Web Stuff
 */
function doGet(e) {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let itemsS = ss.getSheetByName("items")
  if (!itemsS) {
    throw "Error: couldn't find 'items' sheet"
  }
  let priceS = ss.getSheetByName("price_history")
  if (!priceS) {
    throw "Error: couldn't find 'price_history' sheet"
  }
  let homeS = ss.getSheetByName("home")
  if (!homeS) {
    throw "Error: couldn't find 'home' sheet"
  }
  
  let r = {
    "total_value_avg": homeS.getRange("b2").getValue(),
    "total_value_lbin": homeS.getRange("b3").getValue(),
    "total_profit": homeS.getRange("b4").getValue(),
    "total_potential_profit": homeS.getRange("b5").getValue(),
    "total_initial_investment": homeS.getRange("b6").getValue()
  };

  for(let i in r) {
    r[i] = Math.floor(parseInt(r[i])/1e6)
  }
  console.log(r)
  return ContentService.createTextOutput(JSON.stringify(r,null,2)).setMimeType(ContentService.MimeType.JSON)
}
