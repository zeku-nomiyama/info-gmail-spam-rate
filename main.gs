/**
 * Postmaster Tools APIからデータを取得し、そのデータをスプレッドシートに記録し、Slackに通知する関数。
 * 3日前の日付を設定し、その日付を使用してAPIからデータを取得します。
 * 取得したデータは整形され、ドメインごとにスプレッドシートに記録されます。
 * また、スパム率はSlackに通知されます。
 * APIからデータを取得する際にエラーが発生した場合、エラーログをスプレッドシートに記録し、エラーメッセージをSlackに通知します。
 */
function perform() {
  // 最新のデータを取得するための日付設定
  // PostmasterToolsでは3日前のデータが最新であることに注意
  const date = new Date();
  const day = date.getDate();
  date.setDate(day - 3);
  const targetDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd');
  // YYYY/mm/dd形式
  const infoDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');

  try{
      // postMasterToolsからデータ取得
      const apiSettings = setPostmasterToolsAPI();
      const response = UrlFetchApp.fetch(apiSettings.requestUrl, apiSettings.requestOptions);
      const statusCode = response.getResponseCode();
      const responseContent = response.getContent();
      
      // 取得データの整形
      const data = JSON.parse(response);
      const domains = data['domains'];

      //ドメインごとのURLを作成
      const newUrls = []; 
      const length = Object.keys(domains).length;
      for(let i = 0; i < length; i++) {
        let domainName = domains[i].name;
        let url = `https://gmailpostmastertools.googleapis.com/v1/${domainName}/trafficStats`
        newUrls.push(url);
      }

      // ドメインごとにレスポンスを処理
      let successText = [];
      for(let i = 0; i < newUrls.length; i++) {
        const response = UrlFetchApp.fetch(`${newUrls[i]}/${targetDate}`, apiSettings.requestOptions);
        const data = JSON.parse(response);
        const result = formatJsonData(data);

        // スプレッドシートに記録（すべての取得データを記録）
        const domainName = result['ドメイン名'];
        const valueArray = Object.values(result);
        // ドメイン名を日付に置き換える
        valueArray[0] = infoDate;
        recordDataToSpreadsheet(domainName, valueArray);

        // Slackに投稿する文章の作成（スパム率のみ通知）
        successText.push(`${result['ドメイン名']} : ${result['スパム率']}%`);
      }
      // Slackに通知する
      slackPostService(successText,infoDate);

  } catch(e) {
      const errorLog = [e.message];
      errorLog.unshift(infoDate);

      // スプレッドシートに記録（エラーログ）
      recordErrorLogToSpreadsheet(errorLog);

      // Slackに通知する
      const failedText = `スパム率を取得することが出来ませんでした。\n\`\`\`${errorLog[1]}\`\`\``;
      slackPostService(failedText,infoDate);
  }

  // 次回のトリガーを設定
  setTrigger();
}

/**
 * Postmaster Tools APIへのリクエスト設定を生成する関数
 * @return {Object} APIへのリクエスト設定。リクエストオプション（メソッドとヘッダー）とリクエストURLを含むオブジェクトを返す
 */
function setPostmasterToolsAPI() {
  const postmasterToolsService = getPostmasterToolsService_();
  const accessToken = postmasterToolsService.getAccessToken();
  const requestHeaders = {
    'Authorization': 'Bearer ' + accessToken
  }
  const apiSettings = {
    'requestOptions' : {
      'method' : 'get',
      'headers' : requestHeaders
    },
    'requestUrl' : 'https://gmailpostmastertools.googleapis.com/v1/domains'
  }
  return apiSettings;
}

/**
 * Slackへメッセージを投稿する関数
 * @param {any} msg - 投稿するメッセージ
 * @param {string} date - 日付（YYYY/mm/dd）
 */
function slackPostService(msg, date) {
  const token = PropertiesService.getScriptProperties().getProperty("SLACK_TOKEN");
  const slackApp = SlackApp.create(token);
  const channelId = "#bot-test";

  // スプシの名前とURLを取得
  const ss = SpreadsheetApp.openById("1ae1XAmLy9FSv1G1OUn6AgFMQdFpWutgT49IWd8J7OTA");
  const name = ss.getName();
  const url = ss.getUrl(); 

  // メッセージの作成
  const text1 = `<!channel> \n${date}時点のスパム率についてお知らせします。\n`;
  const text2 = `\n詳細については<${url}|${name}>を確認してください。`;
  if (typeof msg === "object") {
    msg = msg.join('\n');
  }
  const message = text1 + msg + text2;

  // メッセージの送信
  slackApp.chatPostMessage(channelId, message);
}

/**
 * スプレッドシートにデータを記録する関数
 * @param {string} sheetName - 書き込みを行うシート名
 * @param {array} data - 書き込みたいデータの配列
 */
function recordDataToSpreadsheet(sheetName, data) {
  // 該当のスプレッドシートを取得
  const ss = SpreadsheetApp.openById("1ae1XAmLy9FSv1G1OUn6AgFMQdFpWutgT49IWd8J7OTA");
  
  // 該当のシートを取得
  const sheet = ss.getSheetByName(sheetName);

  // シートの最終行にデータ書き込み
  sheet.appendRow(data);
}

/**
 * スプレッドシートにエラーログを記録する関数
 * @param {any} errorLog - 書き込みを行うシート名
 */
function recordErrorLogToSpreadsheet(errorLog) {
  // 該当のスプレッドシートを取得
  const ss = SpreadsheetApp.openById("1ae1XAmLy9FSv1G1OUn6AgFMQdFpWutgT49IWd8J7OTA");
  
  // 該当のシートを取得
  const sheet = ss.getSheetByName("エラーログ");

  // シートの最終行にデータ書き込み
  sheet.appendRow(errorLog);
}

/**
 * JSONデータを整形する関数
 * @param {json} data - json data
 * @return {array} result - 連想配列
 */
function formatJsonData(data) {
  // JSONのプロパティ一覧
  const properties = ['name','userReportedSpamRatio','ipReputations','domainReputation','spammyFeedbackLoops','spfSuccessRatio','dkimSuccessRatio','dmarcSuccessRatio','outboundEncryptionRatio','inboundEncryptionRatio','deliveryErrors'];

  // JSONオブジェクト
  const object = {};

  // プロパティの存在チェック
  for (let i = 0; i < properties.length; i++) {
    if(data.hasOwnProperty(properties[i])) {
      object[properties[i]] = data[properties[i]];
    } else {
      object[properties[i]] = null;
    }
  }

  // データの整形
  // ドメイン名
  const domainName = object['name'].match(/\/([^\/]*)\//)[1]; 

  // IPレピュテーション
  let ipReputation = object['ipReputations'];
  let ipReputationCount = {};
  if(ipReputation.length > 0){
    for(let i = 0; i < ipReputation.length; i++){
      if(ipReputation[i].hasOwnProperty('ipCount')){
        ipReputationCount[ipReputation[i]['reputation']] = ipReputation[i]['ipCount'];
      } else {
        ipReputationCount[ipReputation[i]['reputation']] = 0;
      }
    }
  }
  const ipReputationHigh = ipReputationCount['HIGH'];
  const ipReputationMedium = ipReputationCount['MEDIUM'];
  const ipReputationLow = ipReputationCount['LOW'];
  const ipReputationBad = ipReputationCount['BAD'];

  // フィードバックループ
  let feedbackLoop = object['spammyFeedbackLoops'];
  if(feedbackLoop && feedbackLoop.hasOwnProperty('spamRatio')){
    feedbackLoop = feedbackLoop['spamRatio'];
  } else {
    feedbackLoop = null;
  }

  // 配信エラー率
  let errors = {};
  let deliveryErrors = object['deliveryErrors'];
  if(deliveryErrors && deliveryErrors.length > 0){
    for(let i = 0; i < deliveryErrors.length; i++){
      errors[deliveryErrors[i]['errorClass']] = deliveryErrors[i]['errorRatio'];
    }
  }

  let deliveryPermanentErrorRatio = '';
  let deliveryTemporaryErrorRatio = '';
  if(Object.keys(errors).length > 0){ //空オブジェクトかどうか
    deliveryPermanentErrorRatio = errors['PERMANENT_ERROR'];
    deliveryTemporaryErrorRatio = errors['TEMPORARY_ERROR'];
  } else {
    deliveryPermanentErrorRatio = null;
    deliveryTemporaryErrorRatio = null;
  }

  // 整形したjsonデータをオブジェクトにする
  const result = {
    'ドメイン名' : domainName,
    'スパム率' : object['userReportedSpamRatio'] * 100,
    'IPレピュテーションHIGH' : ipReputationHigh,
    'IPレピュテーションMEDIUM' : ipReputationMedium,
    'IPレピュテーションLOW' : ipReputationLow,
    'IPレピュテーションBAD' : ipReputationBad,
    'ドメインレピュテーション' : object['domainReputation'],
    'フィードバックループ' : feedbackLoop * 100,
    'DKIM認証成功率' : object['dkimSuccessRatio'] * 100,
    'SPF認証成功率' : object['spfSuccessRatio'] * 100,
    'DMARC認証成功率' : object['dmarcSuccessRatio'] * 100,
    '受信でのTLS使用率' : object['inboundEncryptionRatio'] * 100,
    '送信でのTLS使用率' : object['outboundEncryptionRatio'] * 100,
    '永続的な配信エラー率' : deliveryPermanentErrorRatio * 100,
    '一時的な配信エラー率' : deliveryTemporaryErrorRatio * 100
  }
 
  return result;
}

/**
 * perform関数の実行トリガーを設定する関数
 */
function setTrigger() {
  let triggers = ScriptApp.getProjectTriggers();
  for(let trigger of triggers){
    let funcName = trigger.getHandlerFunction();
    if(funcName == 'perform'){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  let now = new Date();
  let y = now.getFullYear();
  let m = now.getMonth();
  let d = now.getDate();
  let date = new Date(y, m, d+1, 10, 00);
  ScriptApp.newTrigger('perform').timeBased().at(date).create();
}

// 以下、OAuth2.0の設定に関する関数
function getPostmasterToolsService_() {
  // Create a new service with the given name. The name will be used when
  // persisting the authorized token, so ensure it is unique within the
  // scope of the property store.
  return OAuth2.createService('PostmasterTools')

      // Set the endpoint URLs, which are the same for all Google services.
      .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
      .setTokenUrl('https://accounts.google.com/o/oauth2/token')

      // Set the client ID and secret, from the Google Developers Console.
      .setClientId('792179782882-o7jrub0drplg6qf5sjoe673gcqgnbt5v.apps.googleusercontent.com')
      .setClientSecret('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')

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())

      // Set the scopes to request (space-separated for Google services).
      .setScope('https://www.googleapis.com/auth/postmaster.readonly')

      // Below are Google-specific OAuth2 parameters.

      // Sets the login hint, which will prevent the account chooser screen
      // from being shown to users logged in with multiple accounts.
      .setParam('login_hint', Session.getEffectiveUser().getEmail())

      // Requests offline access.
      .setParam('access_type', 'offline')

      // Consent prompt is required to ensure a refresh token is always
      // returned when requesting offline access.
      .setParam('prompt', 'consent');
}

function createAuthorizationUrl() {
  var postmasterToolsService = getPostmasterToolsService_();
  if (!postmasterToolsService.hasAccess()) {
    var authorizationUrl = postmasterToolsService.getAuthorizationUrl();
    console.log(authorizationUrl);
  } else {
    console.log('認証済み');
  }
}

function authCallback(request) {
  var postmasterToolsService = getPostmasterToolsService_();
  var isAuthorized = postmasterToolsService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}
