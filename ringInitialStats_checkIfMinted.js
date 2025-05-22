const API_URL_BASE = "https://api01.genso.game/api/genso_v2_metadata/";
const ALERT_EMAIL = "bgs.1181.rrdn@gmail.com, 0xpoco@proton.me";
const DISCORD_NOFITY_USERID = "960098451348140072";
const webhookUrl = PropertiesService.getScriptProperties().getProperty("DISCORD_WEBHOOK_URL");

function checkRingMintStatus() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Alert");
  const data = sheet.getDataRange().getValues();

  for (let i = 2; i < data.length; i++) {
    const nftId = data[i][1];
    const monitorFlag = data[i][2];
    const status = data[i][3];

    if (monitorFlag === "ON" && status !== "確認済み") {
      const url = API_URL_BASE + nftId;

      try {
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const code = response.getResponseCode();

        if (code === 200) {
          const json = JSON.parse(response.getContentText());
          const name = json.name;
          const attrs = json.attributes;
          let mintConfirmed = false;
          let inRange = false;

          for (let attr of attrs) {
            const type = attr.trait_type;

            if (type === "hp") inRange = true;

            if (inRange && attr.value !== null && attr.value !== "") {
              mintConfirmed = true;
              break;
            }

            if (type === "exp_get_rate") break;
          }

          if (mintConfirmed) {
            // MailApp.sendEmail(ALERT_EMAIL, `${nftId}がmintされました！(${name})`, `NFT ID ${nftId} がmintされました！\n\n当たりリングつくるぞーー`);
            const message = `<@${DISCORD_NOFITY_USERID}>\n✅ **リングmint通知**\n**NFT ID**: ${nftId} ${name} のmintが確認されました！\n🔗 Create NFTを開く： https://market.genso.game/create-nft/
            `;

            sendDiscordWebhook(message);
            sheet.getRange(i + 1, 4).setValue("確認済み");
            sheet.getRange(i + 1, 3).setValue("OFF"); // 自動OFF

            // 通知送信Eメール
            MailApp.sendEmail(
              ALERT_EMAIL,
              `NFT Mint 完了: ${nftId}`,
              `NFT ID ${nftId} (${name})がMintされました。\n\nCreate NFTを開く： https://market.genso.game/create-nft/`
            );
          }
        }
      } catch (e) {
        Logger.log(`Error with NFT ID ${nftId}: ${e}`);
      }

      // 未mintの場合404を返すケースのコード
      // try {
      //   const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      //   const code = response.getResponseCode();
    }
  }
}

function sendDiscordWebhook(message) {
  const payload = JSON.stringify({
    content: message,
    allowed_mentions: {
      parse: ["users"],
    },
  });

  const options = {
    method: "post",
    contentType: "application/json",
    payload: payload,
    // muteHttpExceptions: true // デバッグ用 - エラーレスポンスをキャプチャ
  };

  // // デバッグ用
  // try {
  //   const response = UrlFetchApp.fetch(webhookUrl, options);
  //   console.log('Response:', response.getContentText()); // 成功時のレスポンス
  // } catch (e) {
  //   console.log('Error:', e.message);
  //   console.log('Response:', e.response ? e.response.getContentText() : 'No response');
  // }

  UrlFetchApp.fetch(webhookUrl, options);
}
