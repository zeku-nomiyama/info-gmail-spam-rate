// oAuth2.0の設定
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

// Postmaster Tools APIへの接続
function setPostmasterToolsAPI() {
  // const requestUrl = "https://gmailpostmastertools.googleapis.com/v1/domains"
  const postmasterToolsService = getPostmasterToolsService_();
  const accessToken = postmasterToolsService.getAccessToken();
  const requestHeaders = {
    'Authorization': 'Bearer ' + accessToken
  }
  // const requestOptions = {
  //   'method' : 'get',
  //   'headers' : requestHeaders
  // }
  const apiSettings = {
    'requestOptions' : {
      'method' : 'get',
      'headers' : requestHeaders
    },
    'requestUrl' : 'https://gmailpostmastertools.googleapis.com/v1/domains'
  }
  return apiSettings;
}

// Slackへの投稿
function slackPostService(msg) {
  const token = PropertiesService.getScriptProperties().getProperty("SLACK_TOKEN");
  const slackApp = SlackApp.create(token);
  const channelId = "#bot-test";
  const message = msg;
  slackApp.chatPostMessage(channelId, message);
}

// 最終的に実行する関数
function perform() {
  // postMasterToolsからデータ取得
  const apiSettings = setPostmasterToolsAPI();
  const response = UrlFetchApp.fetch(apiSettings.requestUrl, apiSettings.requestOptions);

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

  // 最新のデータを取得するための日付設定
  // PostmasterToolsでは2日前のデータが最新であることに注意
  const date = new Date();
  const day = date.getDate();
  date.setDate(day - 2);
  const targetDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd');

  // ドメインごとのレスポンスを取得
  let successText = [];
  for(let i = 0; i < newUrls.length; i++) {
    const response = UrlFetchApp.fetch(`${newUrls[i]}/${targetDate}`, apiSettings.requestOptions);
    const dataByEachDomain = JSON.parse(response);
    // スパム率の取得
    const spamRatio = dataByEachDomain.userReportedSpamRatio;

    // Slackに投稿する文章の作成
    successText.push(`${domains[i].name}の${targetDate}時点のスパム率は${spamRatio * 100}%です。`);
  }

  // Slackへの投稿
  successText = successText.join('\n');
  slackPostService(successText);
}










