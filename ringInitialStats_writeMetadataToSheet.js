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

  // Process only first 3 items for testing
  const TEST_MODE = true;
  const TEST_COUNT = 3;

  // For test mode, directly specify the first 3 token IDs
  if (TEST_MODE) {
    const prefix = "100000000667"; // Life Ring
    const testNftIds = [];
    for (let i = 1; i <= TEST_COUNT; i++) {
      testNftIds.push(`${prefix}${i}`);
    }
    console.log(`Testing with NFT IDs: ${testNftIds.join(", ")}`);

    // For test mode, process only 3 items and exit
    for (const nftId of testNftIds) {
      const mintInfo = getMintInfoForNFT(nftId);
      console.log(`Mint info for ${nftId}:`, mintInfo);
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
 * Get mint info for a specific NFT using Polygon transaction
 * @param {string} nftId - The NFT ID
 * @returns {Object} Object containing mint date and initiator address
 */
function getMintInfoForNFT(nftId) {
  return getDirectMintInfoFromPolygonscan(nftId);
}

function getDirectMintInfoFromPolygonscan(nftId) {
  const contractAddress = "0x0a77f356cf1de1727145e66c92254881ac3da34b"; // GensoKishiOnline-Polygon (ERC-721)
  const gensoCreateNftAddress = "0x801cC71Cad74913f9394806c719C0950b1ec18Ef"; // Genso's wallet address for "createNFT"
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

  // Use tokennfttx API with Genso's address and pagination
  const url = `https://api.polygonscan.com/api?module=account&action=tokennfttx&contractaddress=${contractAddress}&address=${gensoCreateNftAddress}&page=1&offset=10000&sort=desc&apikey=${apiKey}`;

  try {
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());

    console.log(`API Response status: ${data.status}`);
    console.log(`API Response message: ${data.message}`);

    if (data.status === "1" && data.result && data.result.length > 0) {
      console.log(`Number of transactions found: ${data.result.length}`);

      // Filter transactions for the specific tokenId
      const tokenTransactions = data.result.filter(tx => tx.tokenID === nftId);
      console.log(`Number of transactions for tokenId ${nftId}: ${tokenTransactions.length}`);

      if (tokenTransactions.length > 0) {
        console.log(`First transaction for tokenId ${nftId}:`, JSON.stringify(tokenTransactions[0], null, 2));
      }

      // Find the mint transaction (from is zero address)
      const mintTx = tokenTransactions.find(tx => tx.from === "0x0000000000000000000000000000000000000000");

      if (mintTx) {
        console.log(`Found mint transaction: ${mintTx.hash}`);

        // Get transaction details to find the actual initiator
        const txDetailUrl = `https://api.polygonscan.com/api?module=proxy&action=eth_getTransactionByHash&txhash=${mintTx.hash}&apikey=${apiKey}`;
        const txDetailResponse = UrlFetchApp.fetch(txDetailUrl);
        const txDetailData = JSON.parse(txDetailResponse.getContentText());

        console.log(`Transaction details:`, JSON.stringify(txDetailData, null, 2));

        if (txDetailData.result) {
          const timestamp = parseInt(mintTx.timeStamp) * 1000;
          const date = new Date(timestamp);
          const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

          return {
            mint_date: formattedDate,
            initiator_address: txDetailData.result.from, // Actual mint initiator's address
          };
        }
      }
    }

    console.log(`No mint transaction found for NFT ID: ${nftId}`);
    return { mint_date: "", initiator_address: "" };
  } catch (error) {
    console.error(`Error getting transaction details: ${error}`);
    return { mint_date: "", initiator_address: "" };
  }
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
