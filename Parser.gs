/**
 * IMPORTRANGE パーサー
 * 数式から依存関係を抽出するユーティリティ
 */

/**
 * 数式から依存関係を抽出する
 * @param {string} formula - 解析対象の数式
 * @param {string} currentSheetName - 現在のシート名
 * @returns {Object} 依存関係オブジェクト
 */
function parseDependencies(formula, currentSheetName) {
  var dependencies = {
    internal: [],
    externalIds: []
  };

  if (!formula || formula.charAt(0) !== '=') {
    return dependencies;
  }

  // 1. IMPORTRANGE の抽出
  var importRangeRegex = /IMPORTRANGE\s*\(\s*["']([^"']+)["']\s*,\s*["']([^"']+)["']\s*\)/gi;
  var match;
  
  while ((match = importRangeRegex.exec(formula)) !== null) {
    var urlOrId = match[1];
    var rangeRef = match[2];
    
    // URLからID部分を抽出
    var id = urlOrId;
    if (urlOrId.indexOf('/d/') !== -1) {
      var parts = urlOrId.split('/d/');
      if (parts.length > 1) {
        var idPart = parts[1].split('/')[0];
        if (idPart) {
          id = idPart;
        }
      }
    }
    
    // 重複チェック
    var found = false;
    for (var i = 0; i < dependencies.externalIds.length; i++) {
      if (dependencies.externalIds[i].id === id) {
        found = true;
        break;
      }
    }
    
    if (!found) {
      dependencies.externalIds.push({
        id: id,
        range: rangeRef,
        originalUrl: urlOrId
      });
    }
  }

  // 2. 内部シート参照の抽出
  var internalRefRegex = /'([^']+)'![A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?/g;
  
  while ((match = internalRefRegex.exec(formula)) !== null) {
    var sheetName = match[1];
    
    // 現在のシートと異なり、まだ追加されていない場合のみ追加
    if (sheetName !== currentSheetName && dependencies.internal.indexOf(sheetName) === -1) {
      dependencies.internal.push(sheetName);
    }
  }

  // クォートなしの参照も試す
  var simpleRefRegex = /([A-Za-z_\u3040-\u9fff][A-Za-z0-9_\u3040-\u9fff]*)![A-Z]+[0-9]+/g;
  while ((match = simpleRefRegex.exec(formula)) !== null) {
    var sheetName2 = match[1];
    if (sheetName2 !== currentSheetName && dependencies.internal.indexOf(sheetName2) === -1) {
      dependencies.internal.push(sheetName2);
    }
  }

  return dependencies;
}

/**
 * スプレッドシートID/URLからIDを抽出
 * @param {string} input - スプレッドシートIDまたはURL
 * @returns {string} スプレッドシートID
 */
function extractSpreadsheetId(input) {
  if (!input) return null;
  
  input = input.trim();
  
  // URLの場合
  if (input.indexOf('docs.google.com') !== -1) {
    var match = input.match(/\/d\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
  }
  
  // IDそのまま
  return input;
}
