/**
 * 도메인 공유 연락처를 GData m8(Contacts API) XML 방식으로 가져와
 * 시트에 저장하고, 시트 수정사항을 다시 API에 반영하는 예제
 *
 * 📋 사전 준비:
 * 1) Google Cloud Console에서 Contacts API 활성화:
 *    https://console.developers.google.com/apis/api/contacts.googleapis.com/overview?project=<PROJECT_ID>
 * 2) appsscript.json에 아래 OAuth 범위 추가:
 *    "oauthScopes": [
 *      "https://www.googleapis.com/auth/script.external_request",
 *      "https://www.googleapis.com/auth/userinfo.email",
 *      "https://www.googleapis.com/auth/spreadsheets",
 *      "https://www.google.com/m8/feeds"
 *    ]
 * 3) 스크립트 바인딩 후 권한 재승인
 */

// 스프레드시트 설정
var SHEET_ID   = '19-eiz_dSuwemGFuoBl170Dg1omWRQUC4PS5VNcUpbug';
var SHEET_NAME = 'list';

// 네임스페이스
var ATOM_NS     = XmlService.getNamespace('http://www.w3.org/2005/Atom');
var GD_NS       = XmlService.getNamespace('gd',   'http://schemas.google.com/g/2005');
var GCONTACT_NS = XmlService.getNamespace('gContact','http://schemas.google.com/contact/2008');

/**
 * 1) 공유 연락처 조회 및 로깅 강화
 */
function getAllSharedContacts() {
  var domain  = Session.getActiveUser().getEmail().split('@')[1];
  var nextUrl = 'https://www.google.com/m8/feeds/contacts/' + domain + '/full?max-results=1000';
  var contacts = [];

  while (nextUrl) {
    Logger.log('Fetching URL: ' + nextUrl);
    try {
      var resp = UrlFetchApp.fetch(nextUrl, {
        headers:{ 'Authorization':'Bearer '+ScriptApp.getOAuthToken(), 'GData-Version':'3.0' },
        muteHttpExceptions:true
      });
    } catch (e) {
      Logger.log('Fetch exception: ' + e.toString());
      throw e;
    }
    Logger.log('Response code: ' + resp.getResponseCode());
    Logger.log('Response body: ' + resp.getContentText().substring(0,500));
    if (resp.getResponseCode()!==200) {
      throw new Error('Fetch error: HTTP ' + resp.getResponseCode());
    }

    var feed = XmlService.parse(resp.getContentText()).getRootElement();

    feed.getChildren('entry', ATOM_NS).forEach(function(entry) {
      var editLink = entry.getChildren('link', ATOM_NS)
                          .filter(l=>l.getAttribute('rel').getValue()==='edit')[0]
                          .getAttribute('href').getValue();
      var nm = entry.getChild('name', GD_NS) || XmlService.createElement('name', GD_NS);

      var orgs = entry.getChildren('organization', GD_NS).map(function(o){
        var comp  = o.getChildText('orgName', GD_NS)||'';
        var title = o.getChildText('orgTitle',GD_NS)||'';
        return comp + (title? ' ('+title+')':'');
      });

      var addrs = entry.getChildren('structuredPostalAddress', GD_NS).map(function(a){
        var fmt = a.getChildText('formattedAddress', GD_NS);
        if (fmt) return fmt;
        return ['street','city','region','postalCode','country']
          .map(f=>a.getChildText(f, GD_NS)||'')
          .filter(Boolean).join(', ');
      });

      var bdElm = entry.getChild('birthday', GCONTACT_NS);
      var birthday = bdElm && bdElm.getAttribute('when')?
                     bdElm.getAttribute('when').getValue(): '';

      var sites = entry.getChildren('website', GCONTACT_NS).map(function(w){
        return w.getAttribute('href').getValue();
      });

      contacts.push({
        id:         entry.getChildText('id', ATOM_NS),
        editLink:   editLink,
        title:      entry.getChildText('title', ATOM_NS)||'',
        fullName:   nm.getChildText('fullName', GD_NS)||'',
        givenName:  nm.getChildText('givenName',GD_NS)||'',
        familyName: nm.getChildText('familyName',GD_NS)||'',
        emails:     entry.getChildren('email',GD_NS).map(e=>e.getAttribute('address').getValue()),
        phones:     entry.getChildren('phoneNumber',GD_NS).map(p=>p.getText()),
        orgs:       orgs,
        addresses:  addrs,
        birthday:   birthday,
        websites:   sites,
        note:       entry.getChildText('content', ATOM_NS)||'',
        삭제:       ''
      });
    });

    var next = feed.getChildren('link', ATOM_NS)
                   .filter(l=>l.getAttribute('rel').getValue()==='next');
    nextUrl = next.length? next[0].getAttribute('href').getValue(): null;
  }

  writeSheet(contacts);
}

/**
 * 2) 스프레드시트 기록
 */
function writeSheet(contacts) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.clearContents();
  sheet.appendRow([
    'ID','EditLink','Title',
    'Full Name','Given Name','Family Name',
    'Emails','Phones','Organizations',
    'Addresses','Birthday','Websites','Note','삭제'
  ]);
  contacts.forEach(function(c){
    sheet.appendRow([
      c.id, c.editLink, c.title,
      c.fullName, c.givenName, c.familyName,
      c.emails.join('; '),
      c.phones.join('; '),
      c.orgs.join('; '),
      c.addresses.join('; '),
      c.birthday,
      c.websites.join('; '),
      c.note,
      c.삭제
    ]);
  });
  Logger.log('Wrote ' + contacts.length + ' contacts');
}

/**
 * 3) 업데이트 전용 (생략)
 */
function updateSharedContactsFromSheet() {
  // 기존 로직 유지
  Logger.log('updateSharedContactsFromSheet 시작');
  // ...
  Logger.log('updateSharedContactsFromSheet completed');
}

/**
 * 4) 삭제 전용 함수: 로깅 강화
 */
function deleteSharedContactsFromSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found');

  var data    = sheet.getDataRange().getValues();
  var headers = data.shift().map(h=>h.toString().trim().toLowerCase());
  Logger.log('Headers for delete: ' + headers.join(','));
  var idx = {};
  headers.forEach((h,i)=> idx[h]=i);
  var delKey = idx['삭제']!==undefined? '삭제': (idx['delete']!==undefined? 'delete': null);
  if (!delKey) throw new Error("'삭제' or 'delete' column not found");

  data.forEach((r,i)=>{
    var row      = i+2;
    var mark     = (r[idx[delKey]]||'').toString().trim().toLowerCase();
    Logger.log('Row ' + row + ' delete mark: ' + mark);
    if (mark==='y' || mark==='yes') {
      var editLink = r[idx['editlink']];
      Logger.log('Deleting row ' + row + ': ' + editLink);
      try {
        var resp = UrlFetchApp.fetch(editLink, {
          method:'delete',
          headers:{ 'Authorization':'Bearer '+ScriptApp.getOAuthToken(), 'GData-Version':'3.0' },
          muteHttpExceptions:true
        });
        Logger.log('Row ' + row + ' DELETE status: ' + resp.getResponseCode());
        Logger.log('Row ' + row + ' DELETE response: ' + resp.getContentText());
      } catch (e) {
        Logger.log('Row ' + row + ' DELETE exception: ' + e.toString());
      }
    }
  });
  Logger.log('deleteSharedContactsFromSheet completed');
}
