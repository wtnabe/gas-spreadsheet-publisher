/* global SpreadsheetApp, PropertiesService */

/**
 * Library の登録
 *
 * property への必要な値の保存とメニューの追加
 *
 * @param {object} param
 * @param {string} param.targetId コピー先のSpreadsheet ID
 * @param {Array} param.sheetNames
 * @param {boolean} param.append 最後に置くか否か
 * @param {boolean} param.forceReplaceSheet Sheetを強制的に置き換えるか
 */
function register ({ // eslint-disable-line no-unused-vars
  targetId, sheetNames = undefined, append = false, forceReplaceSheet = false
} = {}) {
  if (typeof targetId === 'undefined') {
    lazylog('Missing target Spreadsheet ID with register() function')
  } else {
    setPropTargetSpreadsheet(targetId)
  }
  if (sheetNames && sheetNamesExistOrNot(sheetNames)) setPropSrcSheets(sheetNames)
  setPropFirst(!append)
  setPropForceReplaceSheet(forceReplaceSheet)

  addPublisherMenu()
}

/**
 * @param {Array} sheetNames
 * @returns {boolean}
 */
function sheetNamesExistOrNot (givenSheetNames) {
  const srcSheetNames = SpreadsheetApp.getActive().getSheets().map((sheet) => sheet.getName())

  return givenSheetNames.every((name) => {
    const exist = srcSheetNames.includes(name)
    if (!exist) {
      lazylog(`Sheet ${name} not found in register()`)
    }
    return exist
  })
}

//
// UI
//

/**
 * 専用のメニューを追加
 */
function addPublisherMenu () {
  SpreadsheetApp.getUi()
    .createMenu('Publisher')
    .addItem('Publish', 'SpreadsheetPublisher.copySheets')
    .addItem('Reset properties', 'SpreadsheetPublisher.resetProperties')
    .addToUi()
}

/**
 * Spreadsheet の UI を利用した console.log もどき
 *
 * @param {object} message
 * @returns {void}
 */
function lazylog (message) { // eslint-disable-line no-unused-vars
  SpreadsheetApp.getUi()
    .alert(message)
}

//
// Copy Feature
//

/**
 * 指定のシートをすべて対象の Spreadsheet へコピーする
 *
 * 指定がなかった場合はすべてのシートを対象とする
 *
 * @returns {void}
 */
function copySheets () { // eslint-disable-line no-unused-vars
  const targetSpreadsheet = SpreadsheetApp.openById(propTargetSpreadsheet())
  const lastSheetPosInTarget = targetSpreadsheet.getSheets().length - 1
  const srcSpreadsheet = SpreadsheetApp.getActive()
  const first = propFirst()
  const forceReplaceSheet = propForceReplaceSheet()

  let sheets = propSrcSheets() || srcSpreadsheet.getSheets()
  if (first) sheets = reverseSheets(sheets)

  sheets.forEach((sheet) => {
    copySheetToTarget(
      typeof sheet === 'string' ? srcSpreadsheet.getSheetByName(sheet) : sheet,
      targetSpreadsheet,
      forceReplaceSheet)
    if (first && forceReplaceSheet) {
      // targetSheets は変化しているので都度 getSheets() しないと存在
      // しない sheet へアクセスしようとしてしまう
      targetSpreadsheet.setActiveSheet(targetSpreadsheet.getSheets()[lastSheetPosInTarget])
      targetSpreadsheet.moveActiveSheet(1)
    }
  })
}

/**
 * 指定のシートを対象のSpreadsheetへコピーする
 *
 * 対象の Sheet が存在しない場合は Sheet#copyTo を、すでに存在する場合は Range#copyTo を利用する。
 * "Copy of ..." という名前は元の Sheet と同じになるように rename する
 *
 * @param {Sheet} srcSheet
 * @param {Spreadsheet} targetSpreadsheet Spreadsheet object
 */
function copySheetToTarget (srcSheet, targetSpreadsheet, forceReplaceSheet) {
  const targetSheet = targetSpreadsheet.getSheetByName(srcSheet.getName())

  if (targetSheet) {
    if (forceReplaceSheet) {
      targetSpreadsheet.deleteSheet(targetSheet)
      srcSheet.copyTo(targetSpreadsheet)
      adjustCopiedSheetNameTo(targetSpreadsheet, srcSheet.getName())
    } else {
      const srcRange = srcSheet.getDataRange()
      targetSheet.getRange(srcRange.getA1Notation()).setValues(srcRange.getValues())
    }
  } else {
    srcSheet.copyTo(targetSpreadsheet)
    adjustCopiedSheetNameTo(targetSpreadsheet, srcSheet.getName())
  }
}

/**
 * copyTo した sheet の名前を元の sheet の名前に戻す
 *
 * @param {Spreadsheet} targetSpreadsheet
 * @param {string} srcSheetName
 * @returns {Sheet}
 */
function adjustCopiedSheetNameTo(targetSpreadsheet, srcSheetName) {
  const copiedSheet = targetSpreadsheet.getSheetByName('Copy of ' + srcSheetName)
  return copiedSheet.setName(srcSheetName)
}

/**
 * Sheet の array を reverse() するための function
 *
 * なぜか Sheet の array の reverse() がうまく動かず、そのままの array になってしまうため
 *
 * @param {object} sheets
 * @returns {object}
 */
function reverseSheets (sheets) {
  const result = []

  for (let i = sheets.length - 1; i >= 0; i--) {
    result.push(sheets[i])
  }

  return result
}

//
// Properties
//

/**
 * この Library で利用する property をすべて削除する
 *
 * @returns {void}
 */
function resetProperties () { // eslint-disable-line no-unused-vars
  resetPropTargetSpreadsheet()
  resetPropFirst()
  resetPropSrcSheets()
}

/**
 * コピー先の Spreadsheet ID を property から返す
 *
 * @returns {string}
 */
function propTargetSpreadsheet () {
  return PropertiesService.getUserProperties().getProperty('targetSpreadsheet')
}

/**
 * コピー先の Spreadsheet ID を property へ保存する
 *
 * @param {string} id
 * @returns {PropertyService.Properties}
 */
function setPropTargetSpreadsheet (id) {
  return PropertiesService.getUserProperties().setProperty('targetSpreadsheet', id)
}

/**
 * コピー先の Spreadsheet ID を property から削除する
 *
 * @returns {PropertyService.Properties}
 */
function resetPropTargetSpreadsheet () {
  return PropertiesService.getUserProperties().deleteProperty('targetSpreadsheet')
}

/**
 * @returns {boolean}
 */
function propFirst () {
  return JSON.parse(PropertiesService.getUserProperties().getProperty('placeFirst'))
}

/**
 * @param {boolean}
 * @returns {PropertyService.Properties}
 */
function setPropFirst (first) {
  return PropertiesService.getUserProperties().setProperty('placeFirst', JSON.stringify(first))
}

function resetPropFirst () {
  return PropertiesService.getUserProperties().deleteProperty('placeFirst')
}

/**
 * コピー元の Sheet のリストを property から返す
 *
 * @returns {Array|null}
 */
function propSrcSheets () {
  return JSON.parse(PropertiesService.getUserProperties().getProperty('srcSheets'))
}

/**
 * コピー元の Sheet のリストを property へ保存する
 *
 * @param {Array} srcSheets
 * @returns {PropertyService.Properties}
 */
function setPropSrcSheets (srcSheets) {
  return PropertiesService.getUserProperties().setProperty('srcSheets', JSON.stringify(srcSheets))
}

/**
 * コピー元の Sheet のリストを property から削除する
 *
 * @returns {PropertyService.Properties}
 */
function resetPropSrcSheets () {
  return PropertiesService.getUserProperties().deleteProperty('srcSheets')
}

/**
 * @returns {boolean}
 */
function propForceReplaceSheet () {
  return JSON.parse(PropertiesService.getUserProperties().getProperty('forceReplaceSheet'))
}

/**
 * @param {boolean} forceReplaceSheet
 */
function setPropForceReplaceSheet (forceReplaceSheet) {
  return PropertiesService.getUserProperties().setProperty('forceReplaceSheet', JSON.stringify(forceReplaceSheet))
}
