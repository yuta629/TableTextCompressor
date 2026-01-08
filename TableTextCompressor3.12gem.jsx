/*

テキスト長体調整スクリプト (TableTextCompressor.jsx)

v3.12 - スコープ修正版

【機能概要】

■ セル選択時:
  選択した表セル内のテキストに対して、セル全体で以下の条件で長体(水平比率)を自動調整します。
  
  1. テキストがあふれている場合 (Overflow):
   - [オートON]: あふれが解消されるまで水平比率を下げます。
   - [オートOFF]: 指定行数に収まるまで水平比率を下げます。
   ※ただし、段落数が指定行数より多い場合は処理をスキップします。

  2. テキストがあふれていない場合:
   - 指定した行数に収まるまで水平比率を下げます。
   ※ただし、段落数が指定行数より多い場合は処理をスキップします。

■ TextFrame選択時:
  選択したテキストボックス内のテキストを調整します（セルと同様の処理）
  
  ★追加機能 [先頭文字リセットモード]:
    「ストーリー全段落の先頭文字を100%にする」オプションを使用すると、
    長体調整（圧縮）を行わず、連結されたストーリー全体の全段落の先頭文字を
    強制的に水平比率100%に戻します。

■ テキストダイレクト選択時:
  選択している各段落を個別に処理します（セル内/TextFrame内を自動判定）

【共通仕様】
  - 最小変形率は 40% です。
  - 40%まで圧縮してもあふれが解消しない（または目標行数に収まらない）場合、
    変形率を処理前の状態に戻し、処理をスキップします。

【改善履歴】
v3.12 - スコープ修正版
  - グローバルスコープ汚染を防ぐため、スクリプト全体を即時関数(IIFE)化

v3.11 - 先頭文字リセットモード追加版
  - テキストボックス選択時、長体調整を行わずに先頭文字のみ100%にする単独機能を追加

v3.10 - ロールバック安全性向上版
  - ロールバック（復元）時に再構成(recompose)を行うよう修正

v3.9 - 連結フレーム警告機能追加版
v3.8 - オーバーフロー文字適用漏れ修正版
v3.7 - リファクタリング版
v3.6 - ロールバック機能追加版
v3.5 - パフォーマンス修正版

【動作環境】
Adobe InDesign (macOS / Windows)
ExtendScript (ES3準拠)

*/

#target "InDesign"

(function() { // 即時関数の開始（スコープ汚染防止）

  // ========================================
  // 設定定数
  // ========================================
  var CONFIG = {
    MIN_SCALE: 40,              // 最小水平比率（%）
    STEP: 1,                    // 調整ステップ（%）
    MAX_DEPTH: 20,              // 親要素探索の最大深度
    MAX_TARGET_LINES: 100,      // 目標行数の最大値
    RECOMPOSE_INTERVAL: 5,      // 再計算の間隔（段階数）
    DEBUG_MODE: false           // デバッグモード
  };

  // ========================================
  // メインエントリーポイント
  // ========================================
  try {
    main();
  } catch (e) {
    alert("エラーが発生しました: " + e.message + "\nLine: " + e.line);
  }

  function main() {
    if (app.documents.length === 0) {
      alert("ドキュメントが開かれていません。");
      return;
    }

    if (app.selection.length === 0) {
      alert("セル、テキストボックス、またはテキストを選択してください。");
      return;
    }

    // 選択タイプの判定と対象収集
    var selectionInfo = analyzeSelection(app.selection);
    
    if (selectionInfo.type === "none") {
      alert("有効なセル、テキストボックス、またはテキストが選択されていません。");
      return;
    }

    var params = showDialog(selectionInfo);
    if (!params) return;

    app.doScript(
      function() {
        // ★追加: 先頭文字リセット単独モードの場合
        if (params.resetFirstCharOnly) {
          processResetFirstChar(selectionInfo.targets);
          return;
        }

        // 通常モード
        if (selectionInfo.type === "text") {
          processTextSelection(selectionInfo.targets, params, selectionInfo.containerType);
        } else if (selectionInfo.type === "textframe") {
          processTextFrames(selectionInfo.targets, params);
        } else {
          processCells(selectionInfo.targets, params);
        }
      },
      ScriptLanguage.JAVASCRIPT,
      [],
      UndoModes.ENTIRE_SCRIPT,
      params.resetFirstCharOnly ? "先頭文字100%リセット" : "テキスト長体調整"
    );
  }

  // ========================================
  // 選択範囲解析
  // ========================================

  /**
   * 選択範囲を解析し、タイプと対象を返す
   * @param {Array} selection - app.selection
   * @returns {Object} {type: "cell"|"textframe"|"text"|"none", targets: Array, info: Object}
   */
  function analyzeSelection(selection) {
    var result = {
      type: "none",
      targets: [],
      info: {},
      containerType: null
    };

    if (selection.length === 0) return result;
    
    var item = selection[0];
    
    // 優先順位1: セル選択の判定
    if (item instanceof Cell) {
      var cells = collectTargetCells(selection);
      if (cells.length > 0) {
        result.type = "cell";
        result.targets = cells;
        result.info.cellCount = cells.length;
        result.info.displayName = "表セル";
        return result;
      }
    }
    
    if (item instanceof Table) {
      var cells = collectTargetCells(selection);
      if (cells.length > 0) {
        result.type = "cell";
        result.targets = cells;
        result.info.cellCount = cells.length;
        result.info.displayName = "表セル";
        return result;
      }
    }
    
    // 優先順位2: TextFrame選択
    if (item instanceof TextFrame) {
      var frames = collectTargetTextFrames(selection);
      if (frames.length > 0) {
        result.type = "textframe";
        result.targets = frames;
        result.info.frameCount = frames.length;
        result.info.displayName = "テキストボックス";
        return result;
      }
    }
    
    // 優先順位3: テキストダイレクト選択
    var parentCell = getParentCell(item);
    var parentTextFrame = getParentTextFrame(item);
    
    // 親がセルの場合
    if (parentCell) {
      var paragraphs = collectSelectedParagraphs(item);
      if (paragraphs.length > 0) {
        result.type = "text";
        result.targets = paragraphs;
        result.containerType = "cell";
        result.info.container = parentCell;
        result.info.paragraphCount = paragraphs.length;
        result.info.displayName = "表セル内テキスト";
        return result;
      }
    }
    
    // 親がTextFrameの場合
    if (parentTextFrame) {
      var paragraphs = collectSelectedParagraphs(item);
      if (paragraphs.length > 0) {
        result.type = "text";
        result.targets = paragraphs;
        result.containerType = "textframe";
        result.info.container = parentTextFrame;
        result.info.paragraphCount = paragraphs.length;
        result.info.displayName = "テキストボックス内テキスト";
        return result;
      }
    }

    return result;
  }

  /**
   * 親セルを取得
   * @param {Object} obj - 検索開始オブジェクト
   * @returns {Cell|null} 親セル、見つからない場合はnull
   */
  function getParentCell(obj) {
    var current = obj;
    var depth = 0;
    
    while (current && depth < CONFIG.MAX_DEPTH) {
      if (current instanceof Cell) {
        return current;
      }
      try {
        current = current.parent;
        depth++;
      } catch(e) {
        return null;
      }
    }
    return null;
  }

  /**
   * 親TextFrameを取得（Story対応版）
   * @param {Object} obj - 検索開始オブジェクト
   * @returns {TextFrame|null} 親TextFrame、見つからない場合はnull
   */
  function getParentTextFrame(obj) {
    var current = obj;
    var depth = 0;
    
    while (current && depth < CONFIG.MAX_DEPTH) {
      // セルに到達したら探索中止
      if (current instanceof Cell) {
        return null;
      }
      
      if (current instanceof TextFrame) {
        return current;
      }
      
      // Story経由でTextFrameを取得
      if (current instanceof Story) {
        try {
          if (current.textContainers.length > 0) {
            var container = current.textContainers[0];
            if (container instanceof TextFrame) {
              return container;
            }
          }
        } catch(e) {
          // 続行
        }
      }
      
      try {
        current = current.parent;
        depth++;
      } catch(e) {
        return null;
      }
    }
    return null;
  }

  /**
   * 親コンテナを取得（Cell または TextFrame）
   * @param {Object} obj - 検索開始オブジェクト
   * @returns {Cell|TextFrame|null}
   */
  function getParentContainer(obj) {
    var parentCell = getParentCell(obj);
    if (parentCell) return parentCell;
    
    var parentTextFrame = getParentTextFrame(obj);
    if (parentTextFrame) return parentTextFrame;
    
    return null;
  }

  /**
   * 選択テキストから段落を収集
   * @param {Object} textObj - テキストオブジェクト
   * @returns {Array} Paragraphオブジェクトの配列
   */
  function collectSelectedParagraphs(textObj) {
    var paragraphs = [];
    
    try {
      if (textObj.hasOwnProperty("paragraphs")) {
        var pCount = textObj.paragraphs.length;
        
        if (pCount > 0) {
          for (var i = 0; i < pCount; i++) {
            var para = textObj.paragraphs.item(i);
            if (para && para.isValid) {
              paragraphs.push(para);
            }
          }
          return paragraphs;
        }
      }
      
      var current = textObj;
      var depth = 0;
      while (current && depth < CONFIG.MAX_DEPTH) {
        if (current.hasOwnProperty("paragraphs") && current.paragraphs.length > 0) {
          var pCount = current.paragraphs.length;
          for (var i = 0; i < pCount; i++) {
            var para = current.paragraphs.item(i);
            if (para && para.isValid) {
              paragraphs.push(para);
            }
          }
          return paragraphs;
        }
        
        try {
          current = current.parent;
          depth++;
        } catch(e) {
          break;
        }
      }
      
    } catch(e) {
      if (CONFIG.DEBUG_MODE) {
        alert("段落収集エラー: " + e.message);
      }
    }
    
    return paragraphs;
  }

  /**
   * 処理対象のセルを再帰的に収集する
   * @param {Array} selection - app.selection
   * @returns {Array} Cellオブジェクトの配列
   */
  function collectTargetCells(selection) {
    var cells = [];
    var processed = {};
    
    for (var i = 0; i < selection.length; i++) {
      var item = selection[i];
      
      if (item instanceof Cell) {
        if (item.hasOwnProperty("cells") && item.cells.length > 0) {
          for (var j = 0; j < item.cells.length; j++) {
            var cellId = item.cells[j].id;
            if (!processed[cellId]) {
              cells.push(item.cells[j]);
              processed[cellId] = true;
            }
          }
        } else {
          var cellId = item.id;
          if (!processed[cellId]) {
            cells.push(item);
            processed[cellId] = true;
          }
        }
      }
      else if (item instanceof Table) {
        for (var j = 0; j < item.cells.length; j++) {
          var cellId = item.cells[j].id;
          if (!processed[cellId]) {
            cells.push(item.cells[j]);
            processed[cellId] = true;
          }
        }
      }
      else if (item.parent instanceof Cell) {
        var cellId = item.parent.id;
        if (!processed[cellId]) {
          cells.push(item.parent);
          processed[cellId] = true;
        }
      }
    }
    
    return cells;
  }

  /**
   * 処理対象のTextFrameを収集する
   * @param {Array} selection - app.selection
   * @returns {Array} TextFrameオブジェクトの配列
   */
  function collectTargetTextFrames(selection) {
    var frames = [];
    var processed = {};
    
    for (var i = 0; i < selection.length; i++) {
      var item = selection[i];
      
      if (item instanceof TextFrame) {
        var frameId = item.id;
        if (!processed[frameId]) {
          frames.push(item);
          processed[frameId] = true;
        }
      }
      else if (item.parent instanceof TextFrame) {
        var frameId = item.parent.id;
        if (!processed[frameId]) {
          frames.push(item.parent);
          processed[frameId] = true;
        }
      }
    }
    
    return frames;
  }

  // ========================================
  // ダイアログUI
  // ========================================

  /**
   * 設定ダイアログを表示する
   * @param {Object} selectionInfo - 選択情報
   * @returns {Object|null} 設定オブジェクト、キャンセル時はnull
   */
  function showDialog(selectionInfo) {
    var dlg = new Window("dialog", "テキスト長体調整");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.spacing = 10;
    dlg.margins = 16;

    // --- 対象情報 ---
    var infoGroup = dlg.add("panel", undefined, "対象");
    infoGroup.orientation = "column";
    infoGroup.alignChildren = ["left", "center"];
    
    if (selectionInfo.type === "text") {
      infoGroup.add("statictext", undefined, "選択タイプ: テキストダイレクト選択");
      infoGroup.add("statictext", undefined, "コンテナ: " + selectionInfo.info.displayName);
      infoGroup.add("statictext", undefined, "選択された段落数: " + selectionInfo.info.paragraphCount + " 個");
      var noteText = infoGroup.add("statictext", undefined, "※各段落を個別に処理します");
      noteText.graphics.font = ScriptUI.newFont(noteText.graphics.font.name, "Italic", noteText.graphics.font.size - 1);
    } else if (selectionInfo.type === "textframe") {
      infoGroup.add("statictext", undefined, "選択タイプ: テキストボックス選択");
      infoGroup.add("statictext", undefined, "選択された数: " + selectionInfo.info.frameCount + " 箇所");
    } else {
      infoGroup.add("statictext", undefined, "選択タイプ: 表セル選択");
      infoGroup.add("statictext", undefined, "選択された数: " + selectionInfo.info.cellCount + " 箇所");
    }

    // --- 文字あふれ処理 ---
    var overflowGroup = dlg.add("panel", undefined, "文字あふれ処理");
    overflowGroup.orientation = "column";
    overflowGroup.alignChildren = ["left", "center"];
    var cbAuto = overflowGroup.add("checkbox", undefined, "オート:あふれ解消を優先");
    cbAuto.value = true;
    var helpText = overflowGroup.add("statictext", undefined, "※オフの場合は下の「目標行数」を使用します");
    helpText.graphics.font = ScriptUI.newFont(helpText.graphics.font.name, "Regular", helpText.graphics.font.size - 2);
    helpText.enabled = false;

    // --- 目標行数設定 ---
    var lineGroup = dlg.add("panel", undefined, "目標行数設定");
    lineGroup.orientation = "column";
    lineGroup.alignChildren = ["left", "center"];
    var lineRow = lineGroup.add("group");
    lineRow.orientation = "row";
    lineRow.add("statictext", undefined, "あふれていない場合 / オートOFF時 :");
    var lineInput = lineRow.add("edittext", undefined, "1");
    lineInput.characters = 5;
    lineRow.add("statictext", undefined, "行以内にする");
    
    lineInput.onChanging = function() {
      if (this.text.match(/[^0-9]/)) {
        this.text = this.text.replace(/[^0-9]/g, "");
      }
    };
    
    lineInput.onChange = function() {
      var val = parseInt(this.text, 10);
      if (isNaN(val) || val < 1) {
        this.text = "1";
      } else if (val > CONFIG.MAX_TARGET_LINES) {
        this.text = String(CONFIG.MAX_TARGET_LINES);
      }
    };

    // --- 段落先頭文字の処理 ---
    var charGroup = dlg.add("panel", undefined, "段落先頭文字の処理");
    charGroup.orientation = "column";
    charGroup.alignChildren = ["left", "center"];
    var radioAll = charGroup.add("radiobutton", undefined, "全ての文字を変形する");
    var radioExclude = charGroup.add("radiobutton", undefined, "段落の先頭文字は変形しない");
    radioAll.value = true;

    // ★追加: テキストボックス選択時専用の単独処理オプション
    var cbResetFirst = null;
    if (selectionInfo.type === "textframe") {
      charGroup.add("statictext", undefined, "--------------------------------");
      cbResetFirst = charGroup.add("checkbox", undefined, "ストーリー全段落の先頭文字を100%にする");
      var resetHelp = charGroup.add("statictext", undefined, "※長体調整は行いません。連結フレーム全体に適用されます。");
      resetHelp.graphics.font = ScriptUI.newFont(resetHelp.graphics.font.name, "Italic", resetHelp.graphics.font.size - 2);

      // チェックボックス切り替え時のUI制御
      cbResetFirst.onClick = function() {
        var isResetMode = this.value;
        // 他のコントロールの有効/無効を切り替え
        overflowGroup.enabled = !isResetMode;
        lineGroup.enabled = !isResetMode;
        radioAll.enabled = !isResetMode;
        radioExclude.enabled = !isResetMode;
        
        if (isResetMode) {
          btnOK.text = "リセット実行";
        } else {
          btnOK.text = "実行";
        }
      };
    }

    // --- ボタン ---
    var btnGroup = dlg.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignment = "center";
    var btnCancel = btnGroup.add("button", undefined, "キャンセル", {name: "cancel"});
    var btnOK = btnGroup.add("button", undefined, "実行", {name: "ok"});

    if (dlg.show() == 1) {
      var lines = parseInt(lineInput.text, 10);
      if (isNaN(lines) || lines < 1) lines = 1;
      if (lines > CONFIG.MAX_TARGET_LINES) lines = CONFIG.MAX_TARGET_LINES;
      
      return {
        targetLines: lines,
        excludeFirstChar: radioExclude.value,
        overflowAuto: cbAuto.value,
        resetFirstCharOnly: (cbResetFirst && cbResetFirst.value) ? true : false
      };
    }
    return null;
  }

  // ========================================
  // 共通処理関数 (リファクタリングで追加)
  // ========================================

  /**
   * アイテム配列に対する共通処理ループ
   * @param {Array} items - 処理対象の配列
   * @param {String} uiLabel - UIに表示するラベル（例："段落"）
   * @param {Function} processFunc - 個別処理関数 (item, params) => {changed: bool, skipped: bool}
   * @param {Object} params - 設定パラメータ
   * @returns {Object} 結果 {changed: number, skipped: number, total: number}
   */
  function processItemsCommon(items, uiLabel, processFunc, params) {
    var changedCount = 0;
    var skippedCount = 0;

    var w = new Window("palette", "処理中...");
    var statusText = w.add("statictext", undefined, uiLabel + "を調整しています...");
    var pbar = w.add("progressbar", [0, 0, 300, 20], 0, items.length);
    w.show();

    for (var i = 0; i < items.length; i++) {
      // UI更新頻度を間引く（パフォーマンス向上）
      if (i % 5 === 0) {
        pbar.value = i + 1;
        statusText.text = "処理中: " + (i + 1) + " / " + items.length;
        w.update();
      }

      try {
        var result = processFunc(items[i], params);
        
        // resultがnullまたはundefinedの場合はバリデーションNGとしてスキップ扱い
        if (result) {
          if (result.changed) {
            changedCount++;
          } else if (result.skipped) {
            skippedCount++;
          } else {
            skippedCount++; // changed:false, skipped:false の場合も一応スキップとする
          }
        } else {
          skippedCount++;
        }
      } catch(e) {
        if (CONFIG.DEBUG_MODE) {
          alert(uiLabel + " " + (i + 1) + " のエラー: " + e.message);
        }
        skippedCount++;
      }
    }

    w.close();

    return {
      changed: changedCount,
      skipped: skippedCount,
      total: items.length
    };
  }

  // ========================================
  // 先頭文字リセット処理（新規追加）
  // ========================================

  /**
   * ストーリー全段落の先頭文字を100%にする単独処理
   * @param {Array} frames - 選択されたTextFrame配列
   */
  function processResetFirstChar(frames) {
    var processedStories = {}; // 重複処理防止用
    var totalParagraphs = 0;

    var w = new Window("palette", "処理中...");
    var statusText = w.add("statictext", undefined, "先頭文字をリセットしています...");
    w.show();

    for (var i = 0; i < frames.length; i++) {
      var frame = frames[i];
      var story = frame.parentStory;
      var storyId = story.id;

      // 既に処理済みのストーリーならスキップ（連結フレーム対策）
      if (processedStories[storyId]) continue;
      processedStories[storyId] = true;

      try {
        var paras = story.paragraphs;
        var pLen = paras.length;
        
        for (var p = 0; p < pLen; p++) {
          var para = paras[p];
          if (para.characters.length > 0) {
            para.characters[0].horizontalScale = 100;
            totalParagraphs++;
          }
          
          // UI更新（大量の段落がある場合用）
          if (p % 50 === 0) {
            statusText.text = "処理中: " + (p + 1) + " / " + pLen + " 段落";
            w.update();
          }
        }
        
        // リフローを確定させる
        story.recompose();

      } catch(e) {
        if (CONFIG.DEBUG_MODE) {
          alert("リセット処理エラー: " + e.message);
        }
      }
    }
    
    w.close();

    alert("処理が完了しました。\n" +
          "-----------------------------------\n" +
          "実行内容: 先頭文字100%リセット\n" +
          "対象: 連結ストーリー全体\n" +
          "処理段落数: " + totalParagraphs + " 段落\n" +
          "-----------------------------------\n" +
          "※長体調整は行っていません。");
  }

  // ========================================
  // テキストダイレクト選択時の処理
  // ========================================

  /**
   * テキストダイレクト選択時の処理
   * @param {Array} paragraphs - Paragraphオブジェクトの配列
   * @param {Object} params - 設定パラメータ
   * @param {String} containerType - "cell" または "textframe"
   */
  function processTextSelection(paragraphs, params, containerType) {
    // 共通関数を利用。バリデーションロジックは無名関数でラップして注入。
    var result = processItemsCommon(paragraphs, "段落", function(para, p) {
      if (!para || !para.isValid) return null;
      
      var trimmed = para.contents.replace(/[\r\n\s]/g, "");
      if (trimmed === "") return null;
      
      return adjustSingleParagraph(para, p);
    }, params);

    var containerName = containerType === "cell" ? "セル内" : "テキストボックス内";
    var msg = "処理が完了しました。\n" +
      "-----------------------------------\n" +
      "対象: " + containerName + "テキスト\n" +
      "調整実施: " + result.changed + " 段落\n" +
      "スキップ: " + result.skipped + " 段落\n" +
      "(合計: " + result.total + " 段落)\n" +
      "-----------------------------------\n" +
      "元に戻す場合は ⌘+Z を押してください。";
    alert(msg);
  }

  /**
   * 単一の段落に対する調整ロジック（ロールバック対応）
   * @param {Paragraph} para
   * @param {Object} params
   * @returns {Object} {changed: boolean, skipped: boolean}
   */
  function adjustSingleParagraph(para, params) {
    var startScale = 100;
    
    try {
      if (para.characters.length > 0) {
        startScale = para.characters[0].horizontalScale;
      } else {
        return {changed: false, skipped: false};
      }
    } catch(e) {
      return {changed: false, skipped: false};
    }

    var parentContainer = getParentContainer(para);
    if (!parentContainer) {
      return {changed: false, skipped: false};
    }

    parentContainer.recompose();
    var initialOverflow = isContainerOverflow(parentContainer);
    var currentLines = 0;
    
    try {
      currentLines = para.lines.length;
    } catch(e) {
      return {changed: false, skipped: false};
    }

    var targetMode = determineTargetMode(initialOverflow, currentLines, params);
    if (targetMode === null) {
      return {changed: false, skipped: false};
    }

    var isChanged = false;
    var success = false;
    var firstCharScale = startScale;
    
    if (params.excludeFirstChar && para.characters.length > 0) {
      try {
        firstCharScale = para.characters[0].horizontalScale;
      } catch(e) {
        firstCharScale = 100;
      }
    }

    for (var s = startScale - CONFIG.STEP; s >= CONFIG.MIN_SCALE; s -= CONFIG.STEP) {
      applyScaleToParagraph(para, s, params.excludeFirstChar, firstCharScale);
      isChanged = true;

      var shouldRecompose = ((startScale - s) % CONFIG.RECOMPOSE_INTERVAL === 0) || 
                            (s === CONFIG.MIN_SCALE);
      
      if (shouldRecompose) {
        parentContainer.recompose();
        
        try {
          var nowLines = para.lines.length;
          
          if (targetMode === "lines") {
            if (nowLines <= params.targetLines) {
              success = true;
              break;
            }
          } else {
            var nowOverflow = isContainerOverflow(parentContainer);
            if (!nowOverflow) {
              success = true;
              break;
            }
          }
        } catch(e) {
          break;
        }
      }
    }

    // ロールバック処理
    if (isChanged && !success) {
      applyScaleToParagraph(para, startScale, params.excludeFirstChar, firstCharScale);
      return { changed: false, skipped: true };
    }

    return {
      changed: isChanged,
      skipped: false
    };
  }

  /**
   * モード判定を分離
   * @param {Boolean} isOverflow - オーバーフロー状態
   * @param {Number} currentLines - 現在の行数
   * @param {Object} params - パラメータ
   * @returns {String|null} "overflow"|"lines"|null
   */
  function determineTargetMode(isOverflow, currentLines, params) {
    if (isOverflow) {
      return params.overflowAuto ? "overflow" : "lines";
    } else {
      if (currentLines <= params.targetLines) {
        return null;
      }
      return "lines";
    }
  }

  /**
   * 段落に長体を適用
   * @param {Paragraph} para - 段落オブジェクト
   * @param {Number} scale - 水平比率
   * @param {Boolean} excludeFirst - 先頭文字除外フラグ
   * @param {Number} originalFirstScale - 先頭文字の元の比率
   */
  function applyScaleToParagraph(para, scale, excludeFirst, originalFirstScale) {
    try {
      para.horizontalScale = scale;
      
      if (excludeFirst && para.characters.length > 0) {
        para.characters[0].horizontalScale = originalFirstScale;
      }
    } catch(e) {
      if (CONFIG.DEBUG_MODE) {
        alert("長体適用エラー: " + e.message);
      }
    }
  }

  // ========================================
  // TextFrame選択時の処理
  // ========================================

  /**
   * TextFrame選択時の処理
   * @param {Array} frames - TextFrameオブジェクトの配列
   * @param {Object} params - 設定パラメータ
   */
  function processTextFrames(frames, params) {
    // ★追加: 連結フレームチェック
    var hasLinkedFrame = false;
    for (var i = 0; i < frames.length; i++) {
      // previousTextFrame または nextTextFrame があれば連結されている
      if (frames[i].previousTextFrame !== null || frames[i].nextTextFrame !== null) {
        hasLinkedFrame = true;
        break;
      }
    }

    if (hasLinkedFrame) {
      var confirmMsg = "警告：選択範囲に「連結テキストフレーム」が含まれています。\n" +
                       "処理を実行すると、連結されたストーリー全体（選択していないフレームも含む）が変形対象となります。\n\n" +
                       "処理を続行しますか？";
      // [いいえ] でキャンセル
      if (!confirm(confirmMsg)) {
        return;
      }
    }

    // 共通関数を利用。バリデーションロジックは無名関数でラップして注入。
    var result = processItemsCommon(frames, "テキストボックス", function(frame, p) {
      // texts[0]が存在するかチェック（parentStoryにアクセスする前段として）
      if (frame.texts.length === 0 || frame.texts[0].contents === "") return null;
      return adjustSingleTextFrame(frame, p);
    }, params);

    var msg = "処理が完了しました。\n" +
      "-----------------------------------\n" +
      "対象: テキストボックス\n" +
      "調整実施: " + result.changed + " 箇所\n" +
      "スキップ: " + result.skipped + " 箇所\n" +
      "(合計: " + result.total + " 箇所)\n" +
      "-----------------------------------\n" +
      "元に戻す場合は ⌘+Z を押してください。";
    alert(msg);
  }

  /**
   * 単一のTextFrameに対する調整ロジック（ロールバック対応）
   * @param {TextFrame} frame
   * @param {Object} params
   * @returns {Object} {changed: boolean, skipped: boolean}
   */
  function adjustSingleTextFrame(frame, params) {
    var startScale = 0;
    
    try {
      if (frame.texts.length === 0 || frame.texts[0].characters.length === 0) {
        return {changed: false, skipped: false};
      }
      startScale = frame.texts[0].characters[0].horizontalScale;
    } catch(e) {
      return {changed: false, skipped: false};
    }

    frame.recompose();
    var isOverflow = isContainerOverflow(frame);
    var currentLines = 0;
    var paragraphCount = 0;
    
    try {
      currentLines = frame.texts[0].lines.length;
      paragraphCount = frame.paragraphs.length;
    } catch(e) {
      return {changed: false, skipped: false};
    }

    if (isOverflow) {
      if (params.overflowAuto) {
        // 処理続行
      } else {
        if (paragraphCount > params.targetLines) {
          return {changed: false, skipped: true};
        }
      }
    } else {
      if (currentLines <= params.targetLines) {
        return {changed: false, skipped: false};
      }
      if (paragraphCount > params.targetLines) {
        return {changed: false, skipped: true};
      }
    }

    var targetMode = false;
    
    if (isOverflow) {
      targetMode = !params.overflowAuto;
    } else {
      targetMode = true;
    }

    var appliedScale = startScale;
    var success = false;
    var isChanged = false;

    var firstCharScales = [];
    if (params.excludeFirstChar) {
      for (var p = 0; p < frame.paragraphs.length; p++) {
        var para = frame.paragraphs[p];
        if (para.characters.length > 0) {
          firstCharScales.push(para.characters[0].horizontalScale);
        } else {
          firstCharScales.push(100);
        }
      }
    }

    for (var s = startScale - CONFIG.STEP; s >= CONFIG.MIN_SCALE; s -= CONFIG.STEP) {
      applyScaleToContainer(frame, s, params.excludeFirstChar, firstCharScales);
      isChanged = true;

      // パフォーマンス最適化
      var shouldRecompose = ((startScale - s) % CONFIG.RECOMPOSE_INTERVAL === 0) || 
                            (s === CONFIG.MIN_SCALE);

      if (shouldRecompose) {
        frame.recompose();
        var nowOverflow = isContainerOverflow(frame);
        var nowLines = frame.texts[0].lines.length;

        if (targetMode) {
          // 目標行数モード
          if (nowLines <= params.targetLines) {
            success = true;
            break;
          }
        } else {
          // オーバーフロー解消モード
          if (!nowOverflow) {
            success = true;
            break;
          }
        }
      }
      appliedScale = s;
    }

    // ロールバック処理
    if (isChanged && !success) {
      // ★修正: 第5引数 true を追加して、再構成してから先頭文字を戻す
      applyScaleToContainer(frame, startScale, params.excludeFirstChar, firstCharScales, true);
      return {changed: false, skipped: true};
    }

    return {
      changed: isChanged,
      skipped: false
    };
  }

  // ========================================
  // セル選択時の処理
  // ========================================

  /**
   * セル選択時の処理
   * @param {Array} cells - Cellオブジェクトの配列
   * @param {Object} params - 設定パラメータ
   */
  function processCells(cells, params) {
    // 共通関数を利用。バリデーションロジックは無名関数でラップして注入。
    var result = processItemsCommon(cells, "表セル", function(cell, p) {
      if (cell.texts.length === 0 || cell.texts[0].contents === "") return null;
      return adjustSingleCell(cell, p);
    }, params);

    var msg = "処理が完了しました。\n" +
      "-----------------------------------\n" +
      "対象: 表セル\n" +
      "調整実施: " + result.changed + " 箇所\n" +
      "スキップ: " + result.skipped + " 箇所\n" +
      "(合計: " + result.total + " 箇所)\n" +
      "-----------------------------------\n" +
      "元に戻す場合は ⌘+Z を押してください。";
    alert(msg);
  }

  /**
   * 単一のセルに対する調整ロジック（ロールバック対応）
   * @param {Cell} cell
   * @param {Object} params
   * @returns {Object} {changed: boolean, skipped: boolean}
   */
  function adjustSingleCell(cell, params) {
    var startScale = 0;
    
    try {
      if (cell.texts.length === 0 || cell.texts[0].characters.length === 0) {
        return {changed: false, skipped: false};
      }
      startScale = cell.texts[0].characters[0].horizontalScale;
    } catch(e) {
      return {changed: false, skipped: false};
    }

    cell.recompose();
    var isOverflow = isContainerOverflow(cell);
    var currentLines = 0;
    var paragraphCount = 0;
    
    try {
      currentLines = cell.texts[0].lines.length;
      paragraphCount = cell.paragraphs.length;
    } catch(e) {
      return {changed: false, skipped: false};
    }

    if (isOverflow) {
      if (params.overflowAuto) {
        // 処理続行
      } else {
        if (paragraphCount > params.targetLines) {
          return {changed: false, skipped: true};
        }
      }
    } else {
      if (currentLines <= params.targetLines) {
        return {changed: false, skipped: false};
      }
      if (paragraphCount > params.targetLines) {
        return {changed: false, skipped: true};
      }
    }

    var targetMode = false;
    
    if (isOverflow) {
      targetMode = !params.overflowAuto;
    } else {
      targetMode = true;
    }

    var appliedScale = startScale;
    var success = false;
    var isChanged = false;

    var firstCharScales = [];
    if (params.excludeFirstChar) {
      for (var p = 0; p < cell.paragraphs.length; p++) {
        var para = cell.paragraphs[p];
        if (para.characters.length > 0) {
          firstCharScales.push(para.characters[0].horizontalScale);
        } else {
          firstCharScales.push(100);
        }
      }
    }

    for (var s = startScale - CONFIG.STEP; s >= CONFIG.MIN_SCALE; s -= CONFIG.STEP) {
      applyScaleToContainer(cell, s, params.excludeFirstChar, firstCharScales);
      isChanged = true;

      // パフォーマンス最適化
      var shouldRecompose = ((startScale - s) % CONFIG.RECOMPOSE_INTERVAL === 0) || 
                            (s === CONFIG.MIN_SCALE);

      if (shouldRecompose) {
        cell.recompose();
        var nowOverflow = isContainerOverflow(cell);
        var nowLines = cell.texts[0].lines.length;

        if (targetMode) {
          // 目標行数モード
          if (nowLines <= params.targetLines) {
            success = true;
            break;
          }
        } else {
          // オーバーフロー解消モード
          if (!nowOverflow) {
            success = true;
            break;
          }
        }
      }
      appliedScale = s;
    }

    // ロールバック処理
    if (isChanged && !success) {
      // ★修正: 第5引数 true を追加して、再構成してから先頭文字を戻す
      applyScaleToContainer(cell, startScale, params.excludeFirstChar, firstCharScales, true);
      return {changed: false, skipped: true};
    }

    return {
      changed: isChanged,
      skipped: false
    };
  }

  // ========================================
  // ユーティリティ関数
  // ========================================

  /**
   * コンテナ（Cell/TextFrame）に長体を適用
   * ★v3.10修正: needsRecomposeフラグ追加
   * @param {Cell|TextFrame} container - コンテナオブジェクト
   * @param {Number} scale - 水平比率
   * @param {Boolean} excludeFirst - 先頭文字除外フラグ
   * @param {Array} originalFirstScales - 各段落の先頭文字の元の比率
   * @param {Boolean} [needsRecompose=false] - 適用後に再構成を行うか（ロールバック時用）
   */
  function applyScaleToContainer(container, scale, excludeFirst, originalFirstScales, needsRecompose) {
    try {
      var targetText;
      
      // ★v3.8修正: TextFrameの場合は parentStory を使用
      if (container instanceof TextFrame) {
        targetText = container.parentStory;
      } else {
        // Cellの場合は texts[0] (通常Cell内のテキストは独立しているため)
        targetText = container.texts[0];
      }
      
      // 全体に長体を適用
      targetText.horizontalScale = scale;
      
      // ★v3.10追加: 必要な場合はここで再構成し、段落位置を確定させる
      if (needsRecompose) {
        container.recompose();
      }
      
      if (excludeFirst) {
        // 先頭文字除外処理
        // ※注意: コンテナのparagraphsは可視範囲のみの場合があるが、
        // 溢れている段落の先頭までケアするのは複雑すぎるため、
        // ここでは「コンテナ（frame/cell）から直接取得できる段落」のみを対象とする。
        var paragraphs = container.paragraphs;
        for (var i = 0; i < paragraphs.length; i++) {
          var para = paragraphs[i];
          if (para.characters.length > 0) {
            var original = (originalFirstScales && originalFirstScales.length > i)
              ? originalFirstScales[i]
              : 100;
            para.characters[0].horizontalScale = original;
          }
        }
      }
    } catch(e) {
      if (CONFIG.DEBUG_MODE) {
        alert("コンテナ長体適用エラー: " + e.message);
      }
    }
  }

  /**
   * コンテナ（Cell/TextFrame）のオーバーフロー判定
   * @param {Cell|TextFrame} container - コンテナオブジェクト
   * @returns {Boolean} オーバーフローしている場合true
   */
  function isContainerOverflow(container) {
    return container.overflows;
  }

})(); // 即時関数の終了