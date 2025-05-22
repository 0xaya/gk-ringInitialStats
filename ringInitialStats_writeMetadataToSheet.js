/**
 * Initialize script properties with API keys
 * Run this function once to set up API keys
 */
function setupAPIKeys() {
  const scriptProperties = PropertiesService.getScriptProperties();

  // Set your API keys here
  scriptProperties.setProperties({
    POLYGONSCAN_API_KEY: "MY-POLYGONSCAN-APIKEY",
  });

  Logger.log("API keys setup complete");
}

/**
 * Get contract address from script properties
 * @returns {string} Contract address
 */
function writeMetadataToSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const prefixes = {
    生命: "100000000667",
    // '魔力': '100000000668',
    経験: "100000000669",
    幸運: "100000000670",
    // '腕力': '100000000671',
    知力: "100000000672",
    器用: "100000000673",
    体力: "100000000674",
    速さ: "100000000675",
    精神: "100000000676",
  };

  const keyMap = {
    ID: "nftId",
    Lv: "level",
    HP: "hp",
    MP: "mp",
    腕力: "str",
    体力: "vit",
    速さ: "agi",
    知力: "int",
    器用: "dex",
    精神: "mnd",
    攻撃力: "attack",
    防御力: "defense",
    魔攻: "magic_attack",
    攻撃速度: "atk_spd",
    物CRI値: "physical_cri",
    物CRI倍率: "physical_cri_multi",
    魔CRI値: "magic_cri",
    魔CRI倍率: "magic_cri_multi",
    詠唱速度: "cast_spd",
    防御効率: "def_proficiency",
    ガード: "guard",
    ガード効果: "guard_effect",
    物理: "physical_resist",
    魔: "magic_resist",
    火: "fire_resist",
    風: "wind_resist",
    水: "water_resist",
    土: "earth_resist",
    光: "holy_resist",
    闇: "dark_resist",
    CRI: "critical_resist",
    眠り: "sleep_resist",
    麻痺: "stun_resist",
    毒: "poison_resist",
    沈黙: "silence_resist",
    移動不能: "root_resist",
    移動速度: "snare_resist",
    ドロ率: "item_drop_rate",
    "EXP UP率": "exp_get_rate",
    更新日時: "updated_at",
    Mint日時: "mint_date",
    実行ウォレット: "initiator_address",
  };

  const headers = Object.keys(keyMap);

  // Process each prefix
  for (const [name, prefix] of Object.entries(prefixes)) {
    const startId = 1;
    const endId = 200;

    // Group metadata by name
    const sheetName = `${name}`;
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      for (let row = 2; row <= endId + 1; row++) {
        nftId = `${prefix}${row - 1}`;
        sheet.getRange(row, 1).setValue(nftId);
      }
    }

    const nftIdsToFetch = [];

    // Find the rows where "HP" to "Item Drop Rate" cells are empty
    for (let row = 2; row <= endId + 1; row++) {
      const id = sheet.getRange(row, 1).getValue().toString();
      const startCol = 4; // Column index for "HP"
      const endCol = startCol + 34; // Column index for "Item Drop Rate" + 1
      const range = sheet.getRange(row, startCol, 1, endCol - startCol + 1);
      const values = range.getValues()[0];
      const isAllEmpty = values.every(cell => cell === "");

      if (isAllEmpty) {
        const nftId = `${prefix}${id.slice(12)}`;
        nftIdsToFetch.push(nftId);
      }
    }

    // Fetch metadata for the NFT IDs where "HP" to "Item Drop Rate" cells are empty
    const metadataList = nftIdsToFetch.length > 0 ? retrieveNFTMetadata(nftIdsToFetch) : [];

    // Current datetime for updating records
    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

    // Update the sheet with the new metadata
    for (const metadata of metadataList) {
      // Add update timestamp to metadata
      metadata.updated_at = formattedDate;

      // Attempt to get mint date and initiator address for NFTs that don't have it
      if (!metadata.mint_date || !metadata.initiator_address) {
        const mintInfo = getMintInfoForNFT(metadata.nftId);
        metadata.mint_date = mintInfo.mint_date;
        metadata.initiator_address = mintInfo.initiator_address;
      }

      const existingRow = getExistingRow(sheet, metadata.nftId);
      if (existingRow === -1) {
        // No existing row found, append a new row
        const row = headers.map(header => {
          const key = keyMap[header] || header.toLowerCase().replace(/\s+/g, "_");
          return metadata[key] || "";
        });
        sheet.appendRow(row);
      } else {
        // Existing row found, update the "HP" to "Item Drop Rate" cells and updated_at
        const startCol = 2; // Column index for "HP"
        const endCol = prefix != "100000000669" ? startCol + 36 : startCol + 37; // Column index for "Item Drop Rate" + 1 or "EXP Get Rate" + 1 for EXP Boost Ring

        // Get the last column index (for the update timestamp)
        const lastColumnIndex = headers.length;

        // Update all the values including the timestamp
        const range = sheet.getRange(existingRow, startCol, 1, endCol - startCol + 1);
        const updatedValues = headers.slice(startCol - 1, endCol).map(header => {
          const key = keyMap[header] || header.toLowerCase().replace(/\s+/g, "_");
          return metadata[key] || "";
        });
        range.setValues([updatedValues]);

        // Set the update timestamp separately in the last column
        sheet.getRange(existingRow, lastColumnIndex).setValue(formattedDate);

        // Set the mint date and initiator address if available
        const mintDateColumnIndex = headers.indexOf("Mint日時") + 1;
        const initiatorAddressColumnIndex = headers.indexOf("実行ウォレット") + 1;

        if (mintDateColumnIndex > 0 && metadata.mint_date) {
          sheet.getRange(existingRow, mintDateColumnIndex).setValue(metadata.mint_date);
        }

        if (initiatorAddressColumnIndex > 0 && metadata.initiator_address) {
          sheet.getRange(existingRow, initiatorAddressColumnIndex).setValue(metadata.initiator_address);
        }
      }
    }

    // Update mint info for existing rows that don't have it
    updateMissingMintInfo(sheet, headers);
  }
}

function retrieveNFTMetadata(nftIds) {
  const baseUrl = "https://api01.genso.game/api/genso_v2_metadata/";
  const desiredTraits = [
    "level",
    "hp",
    "mp",
    "attack",
    "defense",
    "magic_attack",
    "atk_spd",
    "str",
    "vit",
    "agi",
    "int",
    "dex",
    "mnd",
    "physical_cri",
    "physical_cri_multi",
    "magic_cri",
    "magic_cri_multi",
    "cast_spd",
    "def_proficiency",
    "guard",
    "guard_effect",
    "physical_resist",
    "magic_resist",
    "fire_resist",
    "wind_resist",
    "water_resist",
    "earth_resist",
    "holy_resist",
    "dark_resist",
    "critical_resist",
    "sleep_resist",
    "stun_resist",
    "poison_resist",
    "silence_resist",
    "root_resist",
    "snare_resist",
    "item_drop_rate",
    "exp_get_rate",
  ];

  const metadataList = [];

  for (const nftId of nftIds) {
    const url = `${baseUrl}${nftId}`;
    const response = UrlFetchApp.fetch(url);

    if (response.getResponseCode() === 404) {
      console.log(`No data found for NFT ID ${nftId}. Skipping.`);
      continue;
    }

    const data = JSON.parse(response.getContentText());

    let metadata = {
      nftId,
      name: data.name,
    };

    // check if all traits are null - if so stop the loop
    let allTraitsNull = true;

    for (const trait of data.attributes) {
      if (desiredTraits.includes(trait.trait_type)) {
        metadata[trait.trait_type] = trait.value;
        if (trait.value !== null && trait.value !== "") {
          allTraitsNull = false;
        }
      }
    }

    if (allTraitsNull) {
      console.log(`All desired traits are null for NFT ID ${nftId}. Stopping loop.`);
      break;
    }

    if (metadata["level"] === "-" || metadata["level"] === null || metadata["level"] === 0) {
      metadata = { ...metadata, level: "-" };
    }
    // console.log(metadata);
    metadataList.push(metadata);
  }

  return metadataList;
}

function getExistingRow(sheet, nftId) {
  const lastRow = sheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    const id = sheet.getRange(row, 1).getValue().toString();
    if (id === nftId) {
      return row;
    }
  }

  return -1;
}

/**
 * Fetch mint info for a specific NFT using Polygon transaction
 * @param {string} nftId - The NFT ID
 * @returns {Object} Object containing mint date and initiator address
 */
function getMintInfoForNFT(nftId) {
  try {
    // API呼び出しを減らすためのキャッシュチェック
    const cacheResult = checkCacheForNFT(nftId);
    if (cacheResult && cacheResult.mint_date && cacheResult.initiator_address) {
      console.log(`Found cached mint info for ${nftId}: ${JSON.stringify(cacheResult)}`);
      return cacheResult;
    }

    console.log(`Starting search for mint info of NFT ID: ${nftId}`);

    // 提供されたトランザクションの例を解析
    if (nftId.startsWith("1000000006")) {
      // リングNFTの場合
      // 例として提供されたトランザクションハッシュを基に検索
      // この例は生命リング (100000000667) に関連すると思われる
      const sampleTxHash = "0x364a2353488a09a6625384c9b0625712afd694efa0c5a3c09f5c437a873d9691";

      // まず、提供されたトランザクションの詳細を確認
      const apiKey = PropertiesService.getScriptProperties().getProperty("POLYGONSCAN_API_KEY");
      const sampleTxDetails = getTransactionDetails(sampleTxHash, apiKey);

      // ヒントとなる情報を出力
      if (sampleTxDetails) {
        console.log(`Sample transaction from: ${sampleTxDetails.from}`);
        console.log(`Sample transaction to: ${sampleTxDetails.to}`);
      }

      // トランザクション入力データからNFT IDを検索
      return searchNFTInTransactionInput(nftId);
    }

    // バックアップ: 一般的な方法でも検索を試みる
    return getDirectMintInfoFromPolygonscan(nftId);
  } catch (error) {
    console.error(`Error in getMintInfoForNFT: ${error}`);
    return { mint_date: "", initiator_address: "" };
  }
}

/**
 * キャッシュから特定のNFT IDの情報を検索
 * @param {string} nftId - 検索するNFT ID
 * @returns {Object|null} キャッシュされた情報または null
 */
function checkCacheForNFT(nftId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const cacheSheetName = "TxCache";
    const cacheSheet = spreadsheet.getSheetByName(cacheSheetName);

    if (!cacheSheet) return null;

    const lastRow = cacheSheet.getLastRow();
    if (lastRow <= 1) return null;

    const idColumn = cacheSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const dateColumn = cacheSheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const addressColumn = cacheSheet.getRange(2, 3, lastRow - 1, 1).getValues();

    for (let i = 0; i < idColumn.length; i++) {
      if (idColumn[i][0] === nftId) {
        return {
          mint_date: dateColumn[i][0],
          initiator_address: addressColumn[i][0],
        };
      }
    }

    return null;
  } catch (error) {
    console.error(`Error checking cache: ${error}`);
    return null;
  }
}

/**
 * リングNFTのサンプルから情報を探索的に取得して一括キャッシュする
 * このメソッドはデバッグと手動実行用
 */
function exploreMintInfoFromSampleTx() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("POLYGONSCAN_API_KEY");
  if (!apiKey) {
    console.error("API key not set");
    return;
  }

  // 生命リングに関連する提供されたサンプルトランザクション
  const sampleTxHash = "0x364a2353488a09a6625384c9b0625712afd694efa0c5a3c09f5c437a873d9691";
  const txDetails = getTransactionDetails(sampleTxHash, apiKey);

  if (txDetails) {
    // このトランザクションの送信者からさらに情報を収集
    const fromAddress = txDetails.from;

    // キャッシュシートを準備
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const cacheSheetName = "ExploredMintInfo";
    let cacheSheet = spreadsheet.getSheetByName(cacheSheetName);

    if (!cacheSheet) {
      cacheSheet = spreadsheet.insertSheet(cacheSheetName);
      cacheSheet.appendRow([
        "Transaction Hash",
        "From Address",
        "To Address",
        "Timestamp",
        "Block Number",
        "Method ID",
        "Function Name",
        "Found Token IDs",
      ]);
    }

    // このアドレスからのトランザクションを取得
    const url = `https://api.polygonscan.com/api?module=account&action=txlist&address=${fromAddress}&startblock=0&endblock=99999999&sort=asc&apikey=${apiKey}`;

    try {
      const response = UrlFetchApp.fetch(url);
      const data = JSON.parse(response.getContentText());

      if (data.status === "1" && data.result && data.result.length > 0) {
        console.log(`Found ${data.result.length} transactions from address ${fromAddress}`);

        // リング関連のコントラクトアドレスへのトランザクションをフィルタリング
        const contractAddress = "0x0a77f356cf1de1727145e66c92254881ac3da34b"; // Genso Kishi Online NFT Contract
        const relevantTxs = data.result.filter(tx => tx.to.toLowerCase() === contractAddress.toLowerCase());

        console.log(`Found ${relevantTxs.length} transactions to Genso NFT contract`);

        // 各トランザクションを処理
        for (const tx of relevantTxs) {
          // 単純なシグネチャ解析（メソッドID = 最初の10文字）
          const methodId = tx.input.substring(0, 10);
          let functionName = "Unknown";

          // 一般的なERC-721関数のメソッドIDをマッピング
          // 実際のメソッドIDは異なる場合があります
          const methodMap = {
            "0x23b872dd": "transferFrom",
            "0x42842e0e": "safeTransferFrom",
            "0xa22cb465": "setApprovalForAll",
            "0x6352211e": "ownerOf",
            "0x95d89b41": "symbol",
            "0x01ffc9a7": "supportsInterface",
            "0x70a08231": "balanceOf",
            "0x40c10f19": "mint",
            "0xd204c45e": "mintTo",
            "0xc87b56dd": "tokenURI",
            "0x4f6ccce7": "tokenByIndex",
            "0x4f558e79": "mintRing", // 仮想的なリング固有の関数
          };

          functionName = methodMap[methodId] || "Unknown";

          // トランザクションの入力からトークンIDを抽出（非常に簡易的）
          // 実際には適切なABIデコードが必要です
          const tokenPattern = /1000000006\d{2}/g;
          const foundTokens = tx.input.match(tokenPattern) || [];

          // 結果をシートに追加
          const timestamp = new Date(parseInt(tx.timeStamp) * 1000);
          const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

          cacheSheet.appendRow([
            tx.hash,
            tx.from,
            tx.to,
            formattedDate,
            tx.blockNumber,
            methodId,
            functionName,
            foundTokens.join(", "),
          ]);
        }
      }
    } catch (error) {
      console.error(`Error exploring mint info: ${error}`);
    }
  }
}
function getMintInfoForNFT(nftId) {
  // バッチ処理で特定のNFT IDに関するトランザクションを直接検索
  return getDirectMintInfoFromPolygonscan(nftId);
}

/**
 * Get transaction details for a specific token ID using Polygonscan API
 * Direct approach to get real mint initiator
 * @param {string} nftId - The NFT ID
 * @returns {Object} Object containing mint date and initiator address
 */
function getDirectMintInfoFromPolygonscan(nftId) {
  const contractAddress = "0x0a77f356cf1de1727145e66c92254881ac3da34b"; // GensoKishiOnline-Polygon (ERC-721)
  const apiKey = PropertiesService.getScriptProperties().getProperty("POLYGONSCAN_API_KEY");

  if (!apiKey) {
    console.error("Polygonscan API key not found in script properties");
    return { mint_date: "", initiator_address: "" };
  }

  // デバッグログ
  console.log(`Searching mint info for NFT ID: ${nftId}`);

  // Get NFT transfer events for this specific token
  // tokenidパラメータを正しく使うために数値のみの部分を抽出
  const tokenIdNumeric = parseInt(nftId.replace(/^[0-9]{12}/, "")); // 先頭12桁を削除して数値のみを取得

  // トークンIDの検索方法を変更 - 数値部分のみで検索
  const url = `https://api.polygonscan.com/api?module=account&action=tokennfttx&contractaddress=${contractAddress}&page=1&offset=100&sort=asc&apikey=${apiKey}`;

  try {
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());

    // デバッグ用にレスポンスをログ出力
    console.log(`API Response status: ${data.status}, Result count: ${data.result ? data.result.length : 0}`);

    if (data.status === "1" && data.result && data.result.length > 0) {
      // 該当するNFTIDを手動でフィルタリング
      // NFT IDの形式を把握し、正確にマッチングするようにする
      console.log(`Looking for tokenID containing: ${tokenIdNumeric}`);

      const matchingTxs = data.result.filter(tx => {
        const txTokenId = tx.tokenID;
        const isMatch = txTokenId.includes(tokenIdNumeric.toString());
        if (isMatch) {
          console.log(`Found matching token: ${txTokenId} for our search ${tokenIdNumeric}`);
        }
        return isMatch;
      });

      console.log(`Found ${matchingTxs.length} matching transactions`);

      if (matchingTxs.length > 0) {
        // 最初のmintトランザクションを見つける（ゼロアドレスからの転送）
        const mintTx = matchingTxs.find(tx => tx.from.toLowerCase() === "0x0000000000000000000000000000000000000000");

        if (mintTx) {
          console.log(`Found mint transaction: ${mintTx.hash}`);

          // トランザクション詳細を取得して実行者を特定
          const txDetails = getTransactionDetails(mintTx.hash, apiKey);

          if (txDetails) {
            const timestamp = parseInt(mintTx.timeStamp) * 1000;
            const date = new Date(timestamp);
            const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

            // 実際のmint実行者
            console.log(`Real initiator: ${txDetails.from}`);

            return {
              mint_date: formattedDate,
              initiator_address: txDetails.from,
            };
          }
        }

        // ゼロアドレスからのトランザクションが見つからない場合は最初のトランザクションを使用
        const firstMatchingTx = matchingTxs[0];
        const timestamp = parseInt(firstMatchingTx.timeStamp) * 1000;
        const date = new Date(timestamp);
        const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

        // トランザクション詳細を取得
        const txDetails = getTransactionDetails(firstMatchingTx.hash, apiKey);
        if (txDetails) {
          return {
            mint_date: formattedDate,
            initiator_address: txDetails.from,
          };
        }

        return {
          mint_date: formattedDate,
          initiator_address: firstMatchingTx.from,
        };
      }
    }

    // 完全なトランザクション履歴を取得して手動で検索する別のアプローチを試す
    return searchMintInfoUsingContractEvents(nftId, apiKey);
  } catch (error) {
    console.error(`Error getting transaction details: ${error}`);
    return { mint_date: "", initiator_address: "" };
  }
}

/**
 * Get transaction details from transaction hash
 * @param {string} txHash - Transaction hash
 * @param {string} apiKey - Polygonscan API key
 * @returns {Object|null} Transaction details or null if error
 */
function getTransactionDetails(txHash, apiKey) {
  const url = `https://api.polygonscan.com/api?module=proxy&action=eth_getTransactionByHash&txhash=${txHash}&apikey=${apiKey}`;

  try {
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());

    if (data.result) {
      return data.result;
    }

    return null;
  } catch (error) {
    console.error(`Error getting transaction details: ${error}`);
    return null;
  }
}

/**
 * 特定のNFT IDのmint情報を検索する代替方法
 * コントラクトのイベントログを使用
 * @param {string} nftId - NFT ID
 * @param {string} apiKey - Polygonscan API key
 * @returns {Object} Mint情報
 */
function searchMintInfoUsingContractEvents(nftId, apiKey) {
  const contractAddress = "0x0a77f356cf1de1727145e66c92254881ac3da34b"; // GensoKishiOnline-Polygon

  // リングの種類に関連するイベントを特定する
  // 元素騎士のリングNFTの場合、特殊なイベントがあるかもしれません
  // NFT IDのプレフィックスを使って特定
  const prefix = nftId.substring(0, 12);

  // 当該リングに対応するイベントトピックを設計
  // 例: keccak256("Transfer(address,address,uint256)")の先頭バイト
  const transferEventTopic = "0xddf252ad1be2c89b69c2b068fc378daa952ba7f163c4a11628f55a4df523b3ef";

  // NFT IDの数値部分を抽出し、16進数に変換（パディング付き）
  const tokenIdNumeric = parseInt(nftId.replace(/^[0-9]{12}/, ""));
  const tokenIdHex = padTo64(tokenIdNumeric.toString(16));

  // ログの取得（複数ページに分けて検索）
  const maxPages = 3;
  let allLogs = [];

  for (let page = 1; page <= maxPages; page++) {
    // フロムアドレスを指定しない場合はすべてのトランザクションを取得
    const url = `https://api.polygonscan.com/api?module=logs&action=getLogs&address=${contractAddress}&topic0=${transferEventTopic}&page=${page}&offset=1000&apikey=${apiKey}`;

    try {
      const response = UrlFetchApp.fetch(url);
      const data = JSON.parse(response.getContentText());

      if (data.status === "1" && data.result && data.result.length > 0) {
        allLogs = allLogs.concat(data.result);
      } else {
        break; // これ以上のログがない場合は終了
      }

      // API制限対策で少し待機
      Utilities.sleep(200);
    } catch (error) {
      console.error(`Error fetching logs page ${page}: ${error}`);
      break;
    }
  }

  console.log(`Found ${allLogs.length} transfer logs total`);

  // トークンIDをトピックから検索するためのパターン設計
  // Transferイベントの場合、通常は最後のトピックがtokenIdになる
  const relevantLogs = allLogs.filter(log => {
    // トピックの内容をチェック（tokenIdがトピックに含まれるかどうか）
    return (
      log.topics &&
      log.topics.length >= 3 &&
      log.topics.some(topic => topic.toLowerCase().includes(tokenIdHex.toLowerCase()))
    );
  });

  console.log(
    `Found ${relevantLogs.length} logs potentially matching our token ID ${tokenIdNumeric} (hex: ${tokenIdHex})`
  );

  if (relevantLogs.length > 0) {
    // 最も古いログ（mint操作に近いもの）を使用
    const oldestLog = relevantLogs.sort((a, b) => parseInt(a.timeStamp, 16) - parseInt(b.timeStamp, 16))[0];

    // トランザクションの詳細を取得
    const txDetails = getTransactionDetails(oldestLog.transactionHash, apiKey);

    if (txDetails) {
      const timestamp = parseInt(oldestLog.timeStamp, 16) * 1000;
      const date = new Date(timestamp);
      const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

      return {
        mint_date: formattedDate,
        initiator_address: txDetails.from,
      };
    } else {
      // トランザクション詳細が取得できない場合はログから時間だけ返す
      const timestamp = parseInt(oldestLog.timeStamp, 16) * 1000;
      const date = new Date(timestamp);
      const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

      return {
        mint_date: formattedDate,
        initiator_address: "",
      };
    }
  }

  // 最後の手段: 空のデータを返す
  return { mint_date: "", initiator_address: "" };
}

/**
 * 16進数文字列を64文字（32バイト）にパディングする
 * @param {string} hex - 16進数文字列（0xプレフィックスなし）
 * @returns {string} 64文字にパディングされた16進数文字列
 */
function padTo64(hex) {
  return hex.padStart(64, "0");
}

/**
 * リング関連のMINTイベントを直接的に検索する
 * GENSOのリング特有のミント関数があるかも
 * @param {string} apiKey - Polygonscan API key
 * @returns {Array} 見つかったトランザクションリスト
 */
function searchRingMintEvents(apiKey) {
  const contractAddress = "0x0a77f356cf1de1727145e66c92254881ac3da34b";
  // リングのmintに特化したイベントやメソッド名があるかもしれません
  // 例えば "MintRing" のようなイベント名
  const possibleEventNames = ["MintRing", "RingMinted", "RingCreated", "ItemMinted"];

  const allEvents = [];

  for (const eventName of possibleEventNames) {
    // イベント名をkeccak256ハッシュ化した値が必要ですが、ここでは仮の値を使用
    // 実際には適切なハッシュ値を使う必要があります
    const eventHash = `0x${eventName.padEnd(64, "0")}`;

    const url = `https://api.polygonscan.com/api?module=logs&action=getLogs&address=${contractAddress}&topic0=${eventHash}&apikey=${apiKey}`;

    try {
      const response = UrlFetchApp.fetch(url);
      const data = JSON.parse(response.getContentText());

      if (data.status === "1" && data.result && data.result.length > 0) {
        console.log(`Found ${data.result.length} logs for event ${eventName}`);
        allEvents.push(...data.result);
      }
    } catch (error) {
      console.error(`Error searching events for ${eventName}: ${error}`);
    }

    // API制限対策で少し待機
    Utilities.sleep(200);
  }

  return allEvents;
}

/**
 * トランザクションバイトコードを解析して特定のリングNFT IDに関連するアクティビティを検索
 *
 * @param {string} nftId - 対象のNFT ID
 * @returns {Object} Mint情報
 */
function searchNFTInTransactionInput(nftId) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("POLYGONSCAN_API_KEY");
  if (!apiKey) return { mint_date: "", initiator_address: "" };

  const contractAddress = "0x0a77f356cf1de1727145e66c92254881ac3da34b";

  // 対象のリングIDの数値部分を16進数に変換
  const tokenIdNumeric = parseInt(nftId.replace(/^[0-9]{12}/, ""));
  const tokenIdHex = tokenIdNumeric.toString(16);
  // 16進数でトークンIDを検索するためのパターン
  const searchPatterns = [tokenIdHex, padTo64(tokenIdHex), tokenIdNumeric.toString()];

  // まず、ゲーム開発者のアドレスなど、リング関連のトランザクションを送信している可能性の高いアドレスからのトランザクションを取得
  const possibleMinterAddresses = [
    "0x364a2353488a09a6625384c9b0625712afd694ef", // サンプルとして提供されたアドレスから取得
    "0x78a3b0a018b9763a67dcfbae7ba7e2e47b9e341f", // 現在取得されている誤ったアドレス
  ];

  for (const address of possibleMinterAddresses) {
    const url = `https://api.polygonscan.com/api?module=account&action=txlist&address=${address}&startblock=0&endblock=99999999&sort=asc&apikey=${apiKey}`;

    try {
      const response = UrlFetchApp.fetch(url);
      const data = JSON.parse(response.getContentText());

      if (data.status === "1" && data.result && data.result.length > 0) {
        console.log(`Checking ${data.result.length} transactions from address ${address}`);

        // トランザクションの入力データで特定のトークンIDパターンを検索
        const matchingTxs = data.result.filter(
          tx =>
            tx.to.toLowerCase() === contractAddress.toLowerCase() &&
            searchPatterns.some(pattern => tx.input.toLowerCase().includes(pattern.toLowerCase()))
        );

        if (matchingTxs.length > 0) {
          console.log(`Found ${matchingTxs.length} transactions potentially related to NFT ${nftId}`);

          // 最も古いトランザクションを使用（おそらくmint）
          const oldestTx = matchingTxs.sort((a, b) => a.timeStamp - b.timeStamp)[0];
          const timestamp = parseInt(oldestTx.timeStamp) * 1000;
          const date = new Date(timestamp);
          const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

          return {
            mint_date: formattedDate,
            initiator_address: oldestTx.from,
          };
        }
      }
    } catch (error) {
      console.error(`Error searching transactions for address ${address}: ${error}`);
    }

    // API制限対策で少し待機
    Utilities.sleep(500);
  }

  // 最終手段：特定のイベントトピックからトランザクションを検索
  return searchMintInfoUsingContractEvents(nftId, apiKey);
}

/**
 * Update missing mint information for existing rows
 * @param {Sheet} sheet - The sheet to update
 * @param {Array} headers - The headers array
 */
function updateMissingMintInfo(sheet, headers) {
  const mintDateColumnIndex = headers.indexOf("Mint日時") + 1;
  const initiatorAddressColumnIndex = headers.indexOf("実行ウォレット") + 1;

  if (mintDateColumnIndex <= 0 && initiatorAddressColumnIndex <= 0) return; // No mint columns found

  const lastRow = sheet.getLastRow();
  const idColumnIndex = 1; // First column contains NFT IDs

  // Process in batches to avoid timeout
  const batchSize = 5; // 少なめのバッチサイズで処理を確実に
  for (let row = 2; row <= lastRow; row += batchSize) {
    const endRow = Math.min(row + batchSize - 1, lastRow);
    const idsRange = sheet.getRange(row, idColumnIndex, endRow - row + 1, 1);
    const ids = idsRange.getValues().flat();

    // Get current values for mint date and address
    const mintDates =
      mintDateColumnIndex > 0
        ? sheet
            .getRange(row, mintDateColumnIndex, endRow - row + 1, 1)
            .getValues()
            .flat()
        : Array(endRow - row + 1).fill("");

    const initiatorAddresses =
      initiatorAddressColumnIndex > 0
        ? sheet
            .getRange(row, initiatorAddressColumnIndex, endRow - row + 1, 1)
            .getValues()
            .flat()
        : Array(endRow - row + 1).fill("");

    let hasChanges = false;
    const updatedMintDates = [...mintDates];
    const updatedInitiatorAddresses = [...initiatorAddresses];

    for (let i = 0; i < ids.length; i++) {
      // If either mint date or initiator address is empty, fetch both
      if (!mintDates[i] || mintDates[i] === "" || !initiatorAddresses[i] || initiatorAddresses[i] === "") {
        const nftId = ids[i];
        if (nftId && nftId !== "") {
          const mintInfo = getMintInfoForNFT(nftId);

          if (mintDateColumnIndex > 0 && mintInfo.mint_date) {
            updatedMintDates[i] = mintInfo.mint_date;
            hasChanges = true;
          }

          if (initiatorAddressColumnIndex > 0 && mintInfo.initiator_address) {
            updatedInitiatorAddresses[i] = mintInfo.initiator_address;
            hasChanges = true;
          }
        }
      }
    }

    // Update only if there are changes
    if (hasChanges) {
      if (mintDateColumnIndex > 0) {
        sheet
          .getRange(row, mintDateColumnIndex, updatedMintDates.length, 1)
          .setValues(updatedMintDates.map(date => [date]));
      }

      if (initiatorAddressColumnIndex > 0) {
        sheet
          .getRange(row, initiatorAddressColumnIndex, updatedInitiatorAddresses.length, 1)
          .setValues(updatedInitiatorAddresses.map(address => [address]));
      }
    }

    // API制限を避けるために待機時間を長めに設定
    Utilities.sleep(2000);
  }
}

/**
 * 特定のトランザクションハッシュからより詳細な情報を取得
 * @param {string} txHash - トランザクションハッシュ
 * @returns {Object} トランザクション詳細情報
 */
function getDetailedTransactionInfo(txHash) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("POLYGONSCAN_API_KEY");

  if (!apiKey) {
    console.error("Polygonscan API key not found in script properties");
    return null;
  }

  // トランザクション実行者取得のためのAPI呼び出し
  const url = `https://api.polygonscan.com/api?module=proxy&action=eth_getTransactionByHash&txhash=${txHash}&apikey=${apiKey}`;

  try {
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());

    if (data.result) {
      return {
        from: data.result.from, // トランザクション実行者
        to: data.result.to, // コントラクトアドレス
        input: data.result.input, // 呼び出したメソッドなどの詳細情報
      };
    }

    return null;
  } catch (error) {
    console.error(`Error fetching transaction details: ${error}`);
    return null;
  }
}

/**
 * バッチ処理でリング特定トランザクションの検索とローカルキャッシュ管理
 * 大量の処理を行う場合に備えてキャッシュ機能を追加
 */
function cacheTransactionData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const cacheSheetName = "TxCache";
  let cacheSheet = spreadsheet.getSheetByName(cacheSheetName);

  if (!cacheSheet) {
    cacheSheet = spreadsheet.insertSheet(cacheSheetName);
    cacheSheet.appendRow(["NFT_ID", "Mint_Date", "Initiator_Address", "Tx_Hash", "Updated_At"]);
  }

  // リングNFTのIDプレフィックスを配列で定義
  const ringPrefixes = [
    "100000000667", // 生命
    "100000000668", // 魔力
    "100000000669", // 経験
    "100000000670", // 幸運
    "100000000671", // 腕力
    "100000000672", // 知力
    "100000000673", // 器用
    "100000000674", // 体力
    "100000000675", // 速さ
    "100000000676", // 精神
  ];

  // 各プレフィックスごとに処理
  for (const prefix of ringPrefixes) {
    // 例として1-10のIDを処理する
    for (let i = 1; i <= 10; i++) {
      const nftId = `${prefix}${i}`;

      // 既存キャッシュをチェック
      const existingRow = findInCache(cacheSheet, nftId);
      if (existingRow > 0) {
        continue; // すでにキャッシュされている場合はスキップ
      }

      // mintデータを取得
      const mintInfo = getMintInfoForNFT(nftId);

      // キャッシュに追加
      if (mintInfo.mint_date && mintInfo.initiator_address) {
        const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
        cacheSheet.appendRow([nftId, mintInfo.mint_date, mintInfo.initiator_address, "", now]);
      }

      // API制限を避けるために待機
      Utilities.sleep(1000);
    }
  }
}

/**
 * キャッシュシートからNFT IDを検索
 * @param {Sheet} sheet - キャッシュシート
 * @param {string} nftId - 検索するNFT ID
 * @returns {number} 行番号または-1（見つからない場合）
 */
function findInCache(sheet, nftId) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1; // ヘッダーのみの場合

  const idColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < idColumn.length; i++) {
    if (idColumn[i][0] === nftId) {
      return i + 2; // 行番号（ヘッダー行 + インデックス + 1）
    }
  }

  return -1;
}
