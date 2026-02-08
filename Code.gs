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
  // サンドボックスモードを明示的に設定
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
      return {
        success: false,
        error: '無効なスプレッドシートID/URLです'
      };
    }

    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheets = ss.getSheets();
    var ssName = ss.getName();

    // ノードとエッジを生成
    var nodes = [];
    var edges = [];
    var externalSpreadsheets = {}; // 外部参照の重複管理

    // 中央ノード（現在のスプレッドシート）
    nodes.push({
      id: spreadsheetId,
      type: 'spreadsheet',
      data: { 
        label: ssName,
        sheetCount: sheets.length
      }
    });

    // 各シートを解析
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName();
      var sheetId = spreadsheetId + '_' + sheetName;
      
      // シートノードを追加
      nodes.push({
        id: sheetId,
        type: 'sheet',
        data: { label: sheetName }
      });

      // スプレッドシート→シートのエッジ
      edges.push({
        id: 'e_' + spreadsheetId + '_' + sheetId,
        source: spreadsheetId,
        target: sheetId,
        type: 'internal'
      });

      // 数式を取得して解析
      var dataRange = sheet.getDataRange();
      var formulas = dataRange.getFormulas();
      var internalRefs = {};
      var externalRefs = {};

      for (var row = 0; row < formulas.length; row++) {
        for (var col = 0; col < formulas[row].length; col++) {
          var formula = formulas[row][col];
          if (formula) {
            var deps = parseDependencies(formula, sheetName);
            
            // 内部参照
            for (var j = 0; j < deps.internal.length; j++) {
              internalRefs[deps.internal[j]] = true;
            }
            
            // 外部参照
            for (var k = 0; k < deps.externalIds.length; k++) {
              var ext = deps.externalIds[k];
              if (!externalRefs[ext.id]) {
                externalRefs[ext.id] = ext;
              }
            }
          }
        }
      }

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
        var ext = externalRefs[extId];
        
        if (!externalSpreadsheets[extId]) {
          // 外部スプレッドシートの情報を取得（可能な場合）
          var extName = extId;
          try {
            var extSs = SpreadsheetApp.openById(extId);
            extName = extSs.getName();
          } catch (e) {
            // アクセス権がない場合はIDのまま
          }
          
          externalSpreadsheets[extId] = extName;
          
          // 外部ノードを追加
          nodes.push({
            id: extId,
            type: 'external',
            data: { 
              label: extName,
              hasAccess: extName !== extId
            }
          });
        }

        edges.push({
          id: 'e_' + extId + '_' + sheetId,
          source: extId,
          target: sheetId,
          type: 'importRange',
          label: ext.range
        });
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
        metadata: {
          spreadsheetName: ssName,
          spreadsheetId: spreadsheetId,
          sheetCount: sheets.length,
          externalCount: externalCount
        }
      }
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * テスト用
 */
function testAnalyze() {
  var result = analyzeSpreadsheet('YOUR_SPREADSHEET_ID_HERE');
  Logger.log(JSON.stringify(result, null, 2));
}
