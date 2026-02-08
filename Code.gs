/**
 * SheetRoots - スプレッドシート依存関係可視化ツール
 * メインエントリーポイント
 */

/**
 * Web アプリのエントリーポイント
 */
function doGet(e) {
  var html = HtmlService.createHtmlOutputFromFile('Index');
  html.setTitle('SheetRoots - 依存関係ビューア');
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  html.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  return html;
}

/**
 * スプレッドシートを解析して依存関係グラフを生成
 * @param {string} spreadsheetInput - スプレッドシートIDまたはURL
 * @returns {Object} ノードとエッジのデータ
 */
function analyzeSpreadsheet(spreadsheetInput) {
  try {
    var spreadsheetId = extractSpreadsheetId(spreadsheetInput);
    if (!spreadsheetId) {
      return { success: false, error: '無効なスプレッドシートID/URLです' };
    }

    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheets = ss.getSheets();
    var ssName = ss.getName();

    var nodes = [];
    var edges = [];
    var externalSpreadsheets = {};
    var allSheetFormulas = {}; // シートごとの数式リスト
    var brokenRefs = []; // 壊れた参照

    // 中央ノード
    nodes.push({
      id: spreadsheetId,
      type: 'spreadsheet',
      data: { 
        label: ssName,
        sheetCount: sheets.length,
        formulas: []
      }
    });

    // 各シートを解析
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName();
      var sheetId = spreadsheetId + '_' + sheetName;
      
      var sheetFormulas = [];
      var internalRefs = {};
      var externalRefs = {};
      var hasError = false;

      // 数式と値を取得
      var dataRange = sheet.getDataRange();
      var formulas = dataRange.getFormulas();
      var values = dataRange.getValues();

      for (var row = 0; row < formulas.length; row++) {
        for (var col = 0; col < formulas[row].length; col++) {
          var formula = formulas[row][col];
          if (formula) {
            var cellAddr = columnToLetter(col + 1) + (row + 1);
            var cellValue = values[row][col];
            var deps = parseDependencies(formula, sheetName);
            
            // 壊れた参照チェック
            var isBroken = false;
            if (typeof cellValue === 'string' && 
                (cellValue.indexOf('#REF!') !== -1 || 
                 cellValue.indexOf('#ERROR!') !== -1 ||
                 cellValue.indexOf('#N/A') !== -1)) {
              isBroken = true;
              hasError = true;
              brokenRefs.push({
                sheet: sheetName,
                cell: cellAddr,
                formula: formula,
                error: cellValue
              });
            }

            // 数式情報を記録
            var refList = [];
            for (var j = 0; j < deps.internal.length; j++) {
              refList.push(deps.internal[j]);
              internalRefs[deps.internal[j]] = true;
            }
            for (var k = 0; k < deps.externalIds.length; k++) {
              var ext = deps.externalIds[k];
              refList.push('EXT:' + ext.id);
              if (!externalRefs[ext.id]) {
                externalRefs[ext.id] = ext;
              }
            }

            sheetFormulas.push({
              cell: cellAddr,
              formula: formula,
              refs: refList,
              isBroken: isBroken,
              error: isBroken ? cellValue : null
            });
          }
        }
      }

      allSheetFormulas[sheetId] = sheetFormulas;

      // シートノードを追加
      nodes.push({
        id: sheetId,
        type: 'sheet',
        data: { 
          label: sheetName,
          formulas: sheetFormulas,
          formulaCount: sheetFormulas.length,
          hasError: hasError,
          referencedBy: [] // 後で計算
        }
      });

      // スプレッドシート→シートのエッジ
      edges.push({
        id: 'e_' + spreadsheetId + '_' + sheetId,
        source: spreadsheetId,
        target: sheetId,
        type: 'internal'
      });

      // 内部シート間のエッジ
      for (var targetSheet in internalRefs) {
        var targetId = spreadsheetId + '_' + targetSheet;
        edges.push({
          id: 'e_' + sheetId + '_' + targetId,
          source: sheetId,
          target: targetId,
          type: 'internalRef'
        });
      }

      // 外部参照のエッジ
      for (var extId in externalRefs) {
        var extRef = externalRefs[extId];
        
        if (!externalSpreadsheets[extId]) {
          var extName = extId;
          var extHasAccess = false;
          try {
            var extSs = SpreadsheetApp.openById(extId);
            extName = extSs.getName();
            extHasAccess = true;
          } catch (e) {
            // アクセス権なし
          }
          
          externalSpreadsheets[extId] = { name: extName, hasAccess: extHasAccess };
          
          nodes.push({
            id: extId,
            type: 'external',
            data: { 
              label: extName,
              hasAccess: extHasAccess,
              hasError: !extHasAccess,
              formulas: [],
              referencedBy: []
            }
          });
        }

        edges.push({
          id: 'e_' + extId + '_' + sheetId,
          source: extId,
          target: sheetId,
          type: 'importRange',
          label: extRef.range
        });
      }
    }

    // 逆依存の計算 (referencedBy)
    for (var e = 0; e < edges.length; e++) {
      var edge = edges[e];
      if (edge.type === 'internalRef' || edge.type === 'importRange') {
        // targetノードを探して、sourceを追加
        for (var n = 0; n < nodes.length; n++) {
          if (nodes[n].id === edge.target) {
            if (!nodes[n].data.referencedBy) {
              nodes[n].data.referencedBy = [];
            }
            var sourceLabel = findNodeLabel(nodes, edge.source);
            if (nodes[n].data.referencedBy.indexOf(sourceLabel) === -1) {
              nodes[n].data.referencedBy.push(sourceLabel);
            }
            break;
          }
        }
      }
    }

    var externalCount = 0;
    for (var key in externalSpreadsheets) {
      externalCount++;
    }

    return {
      success: true,
      data: {
        nodes: nodes,
        edges: edges,
        brokenRefs: brokenRefs,
        metadata: {
          spreadsheetName: ssName,
          spreadsheetId: spreadsheetId,
          sheetCount: sheets.length,
          externalCount: externalCount,
          brokenCount: brokenRefs.length
        }
      }
    };

  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * ノードIDからラベルを検索
 */
function findNodeLabel(nodes, nodeId) {
  for (var i = 0; i < nodes.length; i++) {
    if (nodes[i].id === nodeId) {
      return nodes[i].data.label;
    }
  }
  return nodeId;
}

/**
 * 列番号をアルファベットに変換
 */
function columnToLetter(column) {
  var letter = '';
  while (column > 0) {
    var temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = Math.floor((column - temp - 1) / 26);
  }
  return letter;
}

/**
 * テスト用
 */
function testAnalyze() {
  var result = analyzeSpreadsheet('YOUR_SPREADSHEET_ID_HERE');
  Logger.log(JSON.stringify(result, null, 2));
}
