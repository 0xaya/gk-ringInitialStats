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
    // 魔力: "100000000668",
    経験: "100000000669",
    幸運: "100000000670",
    // 腕力: "100000000671",
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
    Mint日時: "mint_date",
    実行ウォレット: "initiator_address",
  };

  const headers = Object.keys(keyMap);

  // Process only first 3 items for testing
  const TEST_MODE = false;
  const TEST_COUNT = 3;

  // For test mode, directly specify the first 3 token IDs
  if (TEST_MODE) {
    const prefix = "100000000667"; // Life Ring
    const testNftIds = [];
    for (let i = 1; i <= TEST_COUNT; i++) {
      testNftIds.push(`${prefix}${i}`);
    }
    console.log(`Testing with NFT IDs: ${testNftIds.join(", ")}`);

    // Get the sheet for Life Ring
    const sheetName = "生命";
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      // Add all headers including mint info columns
      sheet.appendRow(headers);
    } else {
      // Get current headers
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      // Add missing columns if needed
      if (!currentHeaders.includes("Mint日時")) {
        const lastColumn = sheet.getLastColumn();
        sheet.insertColumnAfter(lastColumn);
        sheet.getRange(1, lastColumn + 1).setValue("Mint日時");
      }
      if (!currentHeaders.includes("実行ウォレット")) {
        const lastColumn = sheet.getLastColumn();
        sheet.insertColumnAfter(lastColumn);
        sheet.getRange(1, lastColumn + 1).setValue("実行ウォレット");
      }
    }

    // For test mode, process only 3 items
    for (const nftId of testNftIds) {
      const mintInfo = getMintInfoForNFT(nftId);
      console.log(`Mint info for ${nftId}:`, mintInfo);

      // Find the row for this NFT ID
      const row = getExistingRow(sheet, nftId);
      if (row > 0) {
        // Get current headers to find correct column indices
        const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const mintDateColumnIndex = currentHeaders.indexOf("Mint日時") + 1;
        const initiatorAddressColumnIndex = currentHeaders.indexOf("実行ウォレット") + 1;

        if (mintDateColumnIndex > 0 && mintInfo.mint_date) {
          sheet.getRange(row, mintDateColumnIndex).setValue(mintInfo.mint_date);
        }

        if (initiatorAddressColumnIndex > 0 && mintInfo.initiator_address) {
          sheet.getRange(row, initiatorAddressColumnIndex).setValue(mintInfo.initiator_address);
        }
      }
    }
    return; // Exit after processing test items
  }

  // Process each prefix
  for (const [name, prefix] of Object.entries(prefixes)) {
    const startId = 1;
    const endId = 200;

    // Group metadata by name
    const sheetName = `${name}`;
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      // Add all headers including mint info columns
      sheet.appendRow(headers);
    } else {
      // Get current headers
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      // Add missing columns if needed
      if (!currentHeaders.includes("Mint日時")) {
        const lastColumn = sheet.getLastColumn();
        sheet.insertColumnAfter(lastColumn);
        sheet.getRange(1, lastColumn + 1).setValue("Mint日時");
      }
      if (!currentHeaders.includes("実行ウォレット")) {
        const lastColumn = sheet.getLastColumn();
        sheet.insertColumnAfter(lastColumn);
        sheet.getRange(1, lastColumn + 1).setValue("実行ウォレット");
      }
    }

    // Get current headers after potential updates
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const mintDateColumnIndex = currentHeaders.indexOf("Mint日時") + 1;
    const initiatorAddressColumnIndex = currentHeaders.indexOf("実行ウォレット") + 1;

    const nftIdsToFetch = [];

    // Process each row
    for (let row = 2; row <= endId + 1; row++) {
      const id = sheet.getRange(row, 1).getValue().toString();
      const nftId = `${prefix}${id.slice(12)}`;

      // Check if the NFT is minted (C to L columns are not empty)
      const startCol = 3; // Column C
      const endCol = 12; // Column L
      const range = sheet.getRange(row, startCol, 1, endCol - startCol + 1);
      const values = range.getValues()[0];
      const isMinted = values.some(cell => cell !== "");

      if (isMinted) {
        // Check if mint info already exists
        const currentMintDate = mintDateColumnIndex > 0 ? sheet.getRange(row, mintDateColumnIndex).getValue() : "";
        const currentInitiatorAddress =
          initiatorAddressColumnIndex > 0 ? sheet.getRange(row, initiatorAddressColumnIndex).getValue() : "";

        // Skip if both mint date and initiator address are already set
        if (currentMintDate && currentInitiatorAddress) {
          // console.log(`Skipping ${nftId} - mint info already exists`);
          continue;
        }

        // Get mint info only if either mint date or initiator address is missing
        const mintInfo = getMintInfoForNFT(nftId);
        console.log(`Mint info for ${nftId}:`, mintInfo);

        if (mintInfo.mint_date && mintInfo.initiator_address) {
          if (mintDateColumnIndex > 0 && !currentMintDate) {
            sheet.getRange(row, mintDateColumnIndex).setValue(mintInfo.mint_date);
          }
          if (initiatorAddressColumnIndex > 0 && !currentInitiatorAddress) {
            sheet.getRange(row, initiatorAddressColumnIndex).setValue(mintInfo.initiator_address);
          }
        }
      } else {
        // Check if the NFT is actually minted by checking mint info
        const mintInfo = getMintInfoForNFT(nftId);
        if (mintInfo.mint_date && mintInfo.initiator_address) {
          // If minted, add to list for metadata fetching
          nftIdsToFetch.push(nftId);
        } else {
          console.log(`Skipping ${nftId} - not minted yet`);
          // Since NFTs are minted sequentially, we can stop checking after finding the first unminted one
          console.log(`Stopping mint check as NFTs are minted sequentially`);
          break;
        }
      }

      // Add delay to avoid API rate limits
      Utilities.sleep(1000);
    }

    // Fetch metadata for unminted NFTs
    if (nftIdsToFetch.length > 0) {
      console.log(`Fetching metadata for ${nftIdsToFetch.length} unminted NFTs`);
      const metadataList = retrieveNFTMetadata(nftIdsToFetch);

      // Update the sheet with the new metadata
      for (const metadata of metadataList) {
        const existingRow = getExistingRow(sheet, metadata.nftId);
        if (existingRow > 0) {
          // Update the "HP" to "Item Drop Rate" cells
          const startCol = 2; // Column index for "HP"
          const endCol = prefix != "100000000669" ? startCol + 36 : startCol + 37; // Column index for "Item Drop Rate" + 1 or "EXP Get Rate" + 1 for EXP Boost Ring

          const range = sheet.getRange(existingRow, startCol, 1, endCol - startCol + 1);
          const updatedValues = headers.slice(startCol - 1, endCol).map(header => {
            const key = keyMap[header] || header.toLowerCase().replace(/\s+/g, "_");
            return metadata[key] || "";
          });
          range.setValues([updatedValues]);

          // Check if mint info is needed
          const currentMintDate =
            mintDateColumnIndex > 0 ? sheet.getRange(existingRow, mintDateColumnIndex).getValue() : "";
          const currentInitiatorAddress =
            initiatorAddressColumnIndex > 0 ? sheet.getRange(existingRow, initiatorAddressColumnIndex).getValue() : "";

          if (!currentMintDate || !currentInitiatorAddress) {
            const mintInfo = getMintInfoForNFT(metadata.nftId);
            console.log(`Mint info for ${metadata.nftId}:`, mintInfo);

            if (mintInfo.mint_date && mintInfo.initiator_address) {
              if (mintDateColumnIndex > 0 && !currentMintDate) {
                sheet.getRange(existingRow, mintDateColumnIndex).setValue(mintInfo.mint_date);
              }
              if (initiatorAddressColumnIndex > 0 && !currentInitiatorAddress) {
                sheet.getRange(existingRow, initiatorAddressColumnIndex).setValue(mintInfo.initiator_address);
              }
            }
          }
        }
      }
    }
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
  const batchSize = 10; // 一度に処理するNFTの数

  // バッチ処理でNFTを処理
  for (let i = 0; i < nftIds.length; i += batchSize) {
    const batch = nftIds.slice(i, i + batchSize);
    console.log(`Processing batch ${i / batchSize + 1} of ${Math.ceil(nftIds.length / batchSize)}`);

    for (const nftId of batch) {
      try {
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

        // check if all traits are null
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
          console.log(`All desired traits are null for NFT ID ${nftId}. No more NFTs to process.`);
          return metadataList; // 連番でミントされるため、ここで処理を終了
        }

        if (metadata["level"] === "-" || metadata["level"] === null || metadata["level"] === 0) {
          metadata = { ...metadata, level: "-" };
        }

        metadataList.push(metadata);
      } catch (error) {
        console.error(`Error processing NFT ID ${nftId}: ${error}`);
        continue;
      }
    }

    // バッチ間でディレイを入れる
    if (i + batchSize < nftIds.length) {
      Utilities.sleep(2000); // 2秒待機
    }
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
 * Get mint info for a specific NFT using Polygon transaction
 * @param {string} nftId - The NFT ID
 * @returns {Object} Object containing mint date and initiator address
 */
function getMintInfoForNFT(nftId) {
  return getDirectMintInfoFromPolygonscan(nftId);
}

function getDirectMintInfoFromPolygonscan(nftId) {
  const contractAddress = "0x0a77f356cf1de1727145e66c92254881ac3da34b"; // GensoKishiOnline-Polygon (ERC-721)
  const operatorAddresses = [
    "0x801cC71Cad74913f9394806c719C0950b1ec18Ef", // First operator wallet
    "0x92D882648Bb00D8c364b7A8302BceA0B1A1754Bb", // Second operator wallet
    "0x9Fe58B38D124771664bDAC8866e42aF12785a5C8", // Current operator wallet
  ];
  const apiKey = PropertiesService.getScriptProperties().getProperty("POLYGONSCAN_API_KEY");

  if (!apiKey) {
    console.error("Polygonscan API key not found in script properties");
    return { mint_date: "", initiator_address: "" };
  }

  // Check and convert nftId type
  if (!nftId) {
    console.error("NFT ID is required");
    return { mint_date: "", initiator_address: "" };
  }

  // Convert to string if number
  nftId = nftId.toString();

  console.log(`Searching mint info for NFT ID: ${nftId}`);

  // Use cached transactions if available
  if (!getDirectMintInfoFromPolygonscan.cachedTransactions) {
    getDirectMintInfoFromPolygonscan.cachedTransactions = {};

    // Fetch transactions for each operator address
    for (const operatorAddress of operatorAddresses) {
      const url = `https://api.polygonscan.com/api?module=account&action=tokennfttx&contractaddress=${contractAddress}&address=${operatorAddress}&page=1&offset=10000&sort=desc&apikey=${apiKey}`;

      try {
        const response = UrlFetchApp.fetch(url);
        const data = JSON.parse(response.getContentText());

        console.log(`API Response status for ${operatorAddress}: ${data.status}`);
        console.log(`API Response message for ${operatorAddress}: ${data.message}`);

        if (data.status === "1" && data.result && data.result.length > 0) {
          console.log(`Number of transactions found for ${operatorAddress}: ${data.result.length}`);
          // Cache the transactions for this operator
          getDirectMintInfoFromPolygonscan.cachedTransactions[operatorAddress] = data.result;
        } else {
          console.log(`No transactions found in API response for ${operatorAddress}`);
        }
      } catch (error) {
        console.error(`Error getting transaction details for ${operatorAddress}: ${error}`);
      }
    }
  }

  // Check transactions from all genso wallets
  for (const operatorAddress of operatorAddresses) {
    const transactions = getDirectMintInfoFromPolygonscan.cachedTransactions[operatorAddress] || [];

    // Filter transactions for the specific tokenId
    const tokenTransactions = transactions.filter(tx => tx.tokenID === nftId);

    // Find the mint transaction (from is zero address)
    const mintTx = tokenTransactions.find(tx => tx.from === "0x0000000000000000000000000000000000000000");

    if (mintTx) {
      // Get transaction details to find the actual initiator
      const txDetailUrl = `https://api.polygonscan.com/api?module=proxy&action=eth_getTransactionByHash&txhash=${mintTx.hash}&apikey=${apiKey}`;
      const txDetailResponse = UrlFetchApp.fetch(txDetailUrl);
      const txDetailData = JSON.parse(txDetailResponse.getContentText());

      if (txDetailData.result) {
        const timestamp = parseInt(mintTx.timeStamp) * 1000;
        const date = new Date(timestamp);
        const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

        return {
          hash: mintTx.hash,
          mint_date: formattedDate,
          initiator_address: txDetailData.result.from, // Actual mint initiator's address
        };
      }
    }
  }

  console.log(`No mint transaction found for NFT ID: ${nftId}`);
  return { mint_date: "", initiator_address: "" };
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
  const batchSize = 5; // Use small batch size for reliable processing
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

    // Set longer wait time to avoid API rate limits
    Utilities.sleep(2000);
  }
}
