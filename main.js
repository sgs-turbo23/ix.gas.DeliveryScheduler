let slack;

function main() {
  slack = new slackNotifier(secret.getSlackId, 'Delivery Schedule');
  try {
    const yesterday = datetimeUtil.getYesterday(new Date());
    // ヨドバシカメラからの配達予定
    setYodobashiDeliverySchedule(yesterday);
    // Amazonからの配達予定
    setAmazonDeliverySchedule(yesterday);
  } catch (error) {
    console.log(error);
    slack.postToSlack(`エラーが発生しました。\n${error}`);
  }
}

function getMessages(query) {
  const threads = GmailApp.search(query);
  return GmailApp.getMessagesForThreads(threads);
}

function registerEvent(title, starttime, endtime, description) {
  var calender = CalendarApp.getCalendarById(secret.getCalendarId);
  var event = calender.createEvent(title, new Date(starttime), new Date(endtime), {description: description});
  event.setColor('4');
}

function registerDayEvent(title, date, description) {
  var calender = CalendarApp.getCalendarById(secret.getCalendarId);
  var event = calender.createAllDayEvent(title, new Date(date), {description: description});
  event.setColor('4');
}

function setAmazonDeliverySchedule(date) {
  // 前日のメールを取得
  const query = `from:auto-confirm@amazon.co.jp after:${datetimeUtil.makeDateString(date)}`;
  // メールの構造はthreads > messsages > thread > messageの構造
  for(const thread of getMessages(query)){
    for(const message of thread){
      // Attachment（添付ファイルを取得）
      const body = message.getBody();
      
      let lines = Parser.data(body)
        .from('お届け予定：')
        .to('</b>')
        .iterate();

      lines.forEach(l => {
        let result = [];
        // </span> <br /> <b> 日曜日, 09/26 
        // </span> <br /> <b> 水曜日, 09/22 08:00 - <br /> 水曜日, 09/22 12:00 
        const matchreg = [...l.matchAll(/[0-9]{2}\/[0-9]{2} [0-9]{2}:[0-9]{2}/g)];
        if (matchreg.length > 1) {
          matchreg.forEach(m => {
            result.push(adjustYear(m[0], date));
          });

          registerEvent('Amazonからの配達', result[0], result[1], 'https://mail.google.com/mail/u/0/#all/' + message.getId());
        } else {
          let match = l.match(/[0-9]{2}\/[0-9]{2} /g);
          let day = adjustYear(match[0].trim(), date);
          registerDayEvent('Amazonからの配達', day, 'https://mail.google.com/mail/u/0/#all/' + message.getId());
        }
      });
      slack.postToSlack(`<@ma.iw>Amazonからの配達予定をGoogle Calendarに保存しました\nhttps://mail.google.com/mail/u/0/#all/` + message.getId());
    }
  }
}

function adjustYear(dateWithoutYear, baseDate) {
  if (dateWithoutYear.match(/[0-9]{2}/g)[0] >= (baseDate.getMonth() + 1)) {
    return `${baseDate.getFullYear()}/${dateWithoutYear}`;
  } else {
    return `${baseDate.getFullYear() + 1}/${dateWithoutYear}`;
  }
}

function setYodobashiDeliverySchedule(date) {
  let schedule = [];

  // 前日のメールを取得
  const query = `from:thanks_gochuumon@yodobashi.com after:${datetimeUtil.makeDateString(date)}`;
  // メールの構造はthreads > messsages > thread > messageの構造
  for(const thread of getMessages(query)){
    for(const message of thread){
      // Attachment（添付ファイルを取得）
      const body = message.getBody();
      const lines = body.split(/\r\n|\n/).filter(v => v.match(".*配達希望日：.*"));

      lines.forEach(l => {
        const line = l.trim().replace('配達希望日：', '').replace('<br>', '');

        const dt = line.split('　');
        // 文字列;　　配達希望日：2021年09月05日　08：00～12：00
        // 年月日と時間を抽出する
        const date = dt[0].replace('年', '/').replace('月', '/').replace('日', '');
        const time = dt[1].split('：').join(':').split('～');
        const result = [`${date} ${time[0]}`, `${date} ${time[1]}`]; 
        if (!schedule.some((x) => x[0] === result[0] && x[1] === result[1])) {
          registerEvent('ヨドバシカメラからの配達', result[0], result[1], 'https://mail.google.com/mail/u/0/#all/' + message.getId());
          schedule.push([`${date} ${time[0]}`, `${date} ${time[1]}`]);
        }
      });
      slack.postToSlack(`<@ma.iw>ヨドバシカメラからの配達予定をGoogle Calendarに保存しました\nhttps://mail.google.com/mail/u/0/#all/` + message.getId());
    }
  }
}