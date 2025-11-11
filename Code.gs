function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SKU Tagger V1')
    .addItem('Start Tagging Process', 'runTagging')
    .addToUi();
}

function writeResultsToSheet(results, thresholds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SPAR (Gampaha)');
  if (!sheet) throw new Error("Could not find 'TaggingSheet' to write results to.");
  if (results.length === 0) return;

  const startRow = 2; 
  const startCol = 2;
  const numRows = results.length;
  const numDataCols = 5;

  const outputData = [];
  const backgroundColors = [];

  for (const res of results) {
    outputData.push(res.output);
    
    if (res.score >= thresholds.HIGH) {
      backgroundColors.push(['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff']);
    } else if (res.score >= thresholds.MEDIUM) {
      backgroundColors.push(['#ffff00', '#ffff00', '#ffff00', '#ffff00', '#ffff00', '#ffff00']);
    } else {
      backgroundColors.push(['#ffcccb', '#ffcccb', '#ffcccb', '#ffcccb', '#ffcccb', '#ffcccb']);
    }
  }

  const clearRange = sheet.getRange(startRow, startCol, sheet.getMaxRows() - 1, numDataCols);
  clearRange.clearContent().setBackground('#ffffff');
  sheet.getRange(startRow, startCol, numRows, numDataCols).setValues(outputData);
  sheet.getRange(startRow, 1, numRows, numDataCols + 1).setBackgrounds(backgroundColors);
  sheet.getRange('B1:F1').setValues([['Generic Keywords', 'Category', 'Basic Type', 'Confidence Score', 'Matched Dictionary SKU']]).setFontWeight('bold');
}

function runTagging() {
  try {
    const thresholds = getConfig();
    const dictionary = getDictionaryData();
    const skusToTag = getTaggingSKUs();
    if (dictionary.length === 0) throw new Error("EMPTY DICTIONARY: Your 'Dictionary' sheet is empty.");
    if (skusToTag.length === 0) throw new Error("NO SKUs: No SKUs found in 'TaggingSheet'.");
    
    const finalResults = [];
    let highConfidenceCount = 0;
    let mediumConfidenceCount = 0;
    
    for (const sku of skusToTag) {
      const skuTokens = preprocessText(sku);
      const { match, score } = findBestMatch(skuTokens, dictionary);
      
      let rowData;
      
      if (score >= thresholds.MEDIUM) {
        if (score >= thresholds.HIGH) highConfidenceCount++;
        else mediumConfidenceCount++;

        const { originalData } = match;
        const combinedKeywords = [originalData.visibleKeyword, originalData.notVisibleKeyword].filter(Boolean).join(', ');
        
        rowData = [
          combinedKeywords,
          originalData.category,
          originalData.basicType,
          (score * 100).toFixed(2) + '%',
          originalData.skuName
        ];
      } else {
        rowData = ['', '', '', (score * 100).toFixed(2) + '%', 'No confident match found'];
      }
      finalResults.push({ output: rowData, score: score });
    }
    
    writeResultsToSheet(finalResults, thresholds);

    const totalTagged = highConfidenceCount + mediumConfidenceCount;
    const summary = `✅ Tagging Complete! \n\n` +
      `Total SKUs Processed: ${skusToTag.length}\n` +
      `-----------------------------------\n` +
      `▪️ ${highConfidenceCount} tagged with high confidence.\n` +
      `▪️ ${mediumConfidenceCount} tagged for review (yellow).\n` +
      `▪️ ${skusToTag.length - totalTagged} could not be matched (red).\n\n` +
      `Your sheet has been updated.`;
      
    SpreadsheetApp.getUi().alert(summary);

  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ An error occurred:\n\n${error.message}\n\nPlease check your sheet names and data formats.`);
    Logger.log(error);
  }
}

function getConfig() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('Config');
    if (!configSheet) {
      Logger.log("Config sheet not found, using default thresholds.");
      return { HIGH: 0.98, MEDIUM: 0.70 };
    }
    const high = parseFloat(configSheet.getRange('B1').getValue());
    const medium = parseFloat(configSheet.getRange('B2').getValue());
    return { HIGH: high || 0.98, MEDIUM: medium || 0.70 };
}

function preprocessText(text) {
  if (!text || typeof text !== 'string') return new Set();
  let cleanText = text.toLowerCase().replace(/\|/g, ' ').replace(/\(.*?\)/g, '').replace(/[\d\.]+(kg|g|ml|l)\b/g, '').replace(/[^a-z0-9\s]/g, '').trim().replace(/\s+/g, ' ');
  const tokens = cleanText.split(' ');
  const stemmedTokens = tokens.map(token => token.length > 3 && token.endsWith('s') ? token.slice(0, -1) : token);
  return new Set(stemmedTokens.filter(token => token.length > 0));
}

function sorensenDiceScore(setA, setB) {
  const setASize = setA.size;
  const setBSize = setB.size;
  if (setASize === 0 && setBSize === 0) return 1;
  if (setASize === 0 || setBSize === 0) return 0;
  let intersectionCount = 0;
  const unmatchedB = new Set(setB);
  for (const tokenA of setA) {
    let bestFuzzyMatch = null;
    let minDistance = 2;
    if (unmatchedB.has(tokenA)) {
      bestFuzzyMatch = tokenA;
    } else {
      for (const tokenB of unmatchedB) {
        const distance = levenshteinDistance(tokenA, tokenB);
        if (tokenA.length >= 4 && distance < minDistance) {
          minDistance = distance;
          bestFuzzyMatch = tokenB;
        }
      }
    }
    if (bestFuzzyMatch) {
      intersectionCount++;
      unmatchedB.delete(bestFuzzyMatch);
    }
  }
  return (2 * intersectionCount) / (setASize + setBSize);
}

function levenshteinDistance(s1, s2) {
  s1 = s1.toLowerCase(); s2 = s2.toLowerCase(); const costs = [];
  for (let i = 0; i <= s1.length; i++) {
    let lastValue = i;
    for (let j = 0; j <= s2.length; j++) {
      if (i === 0) costs[j] = j;
      else if (j > 0) {
        let newValue = costs[j - 1];
        if (s1.charAt(i - 1) !== s2.charAt(j - 1)) newValue = Math.min(newValue, lastValue, costs[j]) + 1;
        costs[j - 1] = lastValue; lastValue = newValue;
      }
    }
    if (i > 0) costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}

function findBestMatch(skuTokens, dictionary) {
  let bestScore = -1; let bestMatch = null;
  for (const entry of dictionary) {
    const currentScore = sorensenDiceScore(skuTokens, entry.tokens);
    if (currentScore > bestScore) {
      bestScore = currentScore;
      bestMatch = entry;
    }
  }
  return { match: bestMatch, score: bestScore };
}

function getDictionaryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dictionarySheet = ss.getSheetByName('Dictionary');
  if (!dictionarySheet) throw new Error('SHEET NOT FOUND: "Dictionary" tab is missing.');
  const dataRange = dictionarySheet.getRange(2, 1, dictionarySheet.getLastRow() - 1, 5).getValues();
  return dataRange.map(row => ({ originalData: { skuName: row[0], visibleKeyword: row[1], notVisibleKeyword: row[2], category: row[3], basicType: row[4] }, tokens: preprocessText(row[0]) })).filter(e => e.originalData.skuName && e.originalData.skuName.trim() !== '');
}

function getTaggingSKUs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taggingSheet = ss.getSheetByName('SPAR (Gampaha)');
  if (!taggingSheet) throw new Error('SHEET NOT FOUND: "TaggingSheet" tab is missing.');
  if (taggingSheet.getLastRow() <= 1) return [];
  return taggingSheet.getRange(2, 1, taggingSheet.getLastRow() - 1, 1).getValues().flat().filter(sku => sku && sku.trim() !== '');
}