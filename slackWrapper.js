class slackWrapper {
  constructor(url, name) {
    this.postUrl = url;
    this.username = name;
  }

  postToSlack(message) {
    var jsonData =
    {
      "username" : this.username,
      "text" : message
    };
    var payload = JSON.stringify(jsonData);

    var options =
    {
      "method" : "post",
      "contentType" : "application/json",
      "payload" : payload
    };

    UrlFetchApp.fetch(this.postUrl, options);
  }
}