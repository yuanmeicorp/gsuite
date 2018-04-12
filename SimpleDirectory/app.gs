/**
 * Yuan Mei Corp.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Add '*' to first column if there is any change
 * 如果資料有變動，在第一欄內填入 '*'
 */
function onEdit(e) {
  if (DirectorySheetName == e.range.getSheet().getName()) {
    r = e.range.getRow();
    c = e.range.getColumn();
    h = e.range.getHeight();
    if (r>1) {
      e.range.getSheet().getRange(r,1,h,1).setValue('*');
    }
  }
}

/**
 * Add menu to Google Sheets UI
 * 把選單加入 Google Sheets 用戶介面
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(AppTitle)
      .addItem(MenuDownload, 'dirDownload')
      .addItem(MenuUpload, 'dirUpload')
      .addSeparator()
      .addItem(MenuSetup, 'initDomain')
      .addToUi();
}

var timeZone;

/**
 * Return the sheet. Set up one if it doesn't exist.
 * 回傳工作表。如果無工作表，設置一個。
 * @returns {Sheet}
 */
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DirectorySheetName);
  if (sheet == null) {
    sheet = ss.insertSheet(DirectorySheetName, 1);
    sheet.getRange(1/*row*/,2/*col*/, 1, headers.length).setValues([headers]);
    sheet.getRange('A:D').protect().setWarningOnly(true);
  } else {
    // clear all sheet data first
    // 先把工作表資料全部刪除
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).clear();
    }
  }
  return sheet;
}

function initDomain() {
  var docProperties =  PropertiesService.getDocumentProperties();
  var ui = SpreadsheetApp.getUi();
  var gsDomain = docProperties.getProperty('SimpleDirectory.domain')
  var response = ui.prompt(AppTitle, MsgInitDomain.replace('$Domain$', gsDomain), ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    docProperties.setProperty('SimpleDirectory.domain', response.getResponseText());
  }
}

/**
 * Return 'Never logged in' or date in 'yyyy-MM-dd hh:mm' format
 * 回傳 '從未登入' 或是 'yyyy-MM-dd hh:mm' 格式的日期
 * @param {Date} loginTime  登入時間
 * @returns {String}
 */
function formatLoginTime(loginTime) {
  d = new Date(loginTime);
  return d.getTime()?'\''+Utilities.formatDate(d, timeZone, 'yyyy-MM-dd hh:mm'):MsgNeverLoggedIn;
}

/**
 * Convert user info to serialized array
 * 把使用者資料轉成一維陣列
 * @param {Object} user 使用者
 * @returns {...String}
 */
function serializeUserInfo(user) {
  var title = '';
  var department = '';
  
  if (user.organizations) { 
    title = user.organizations[0].title||'';
    department = user.organizations[0].department||'';
  }
  
  var lastLoginTime = formatLoginTime(user.lastLoginTime);
  
  return [
          user.name.fullName,                      //        0 全名
          user.primaryEmail,                       //        1 主要電子郵箱
          lastLoginTime,                           //        2 最後一次登入時間
          getEmployId(user),                       //  0     3 員工編號
          title,                                   //  1     4 職稱
          department,                              //  2     5 部門
          getManager(user)                         //  3     6 主管
        ]
        .concat( getAssistants(user, 2) )          // 4,5  7,8 助理
        .concat(
            [
              getPhones(user, 'work' , 1)[0],      //  6     9 公司電話
              getPhones(user, 'mobile', 1)[0],     //  7    10 行動電話
              getIMs(user, 'skype', 1)[0],         //  8    11 Skype
              getIMs(user, 'line', 1)[0],          //  9    12 Line
              user.notes?user.notes.value:'',      // 10    13 備註
              user.includeInGlobalAddressList,     // 11    14 公開否
              user.orgUnitPath                     // 12    15 組織路徑
            ]);
}

/**
 * Update user info
 * 更新使用者資料
 * @param {String} userKey 使用者帳號 (user@site.com)
 * @param {...String} info 使用者資料
 */
function updateUser(userKey, info) {
  var user = AdminDirectory.Users.get(userKey);
  setPropListValues(user,'externalIds', 'type', 'organization', 'value', info[0],true);
  
  var ttl = String(info[1]).trim();
  var dpt = String(info[2]).trim();
  if (user.organizations) { 
    user.organizations[0].title = ttl;
    user.organizations[0].department = dpt;
  } else {
    if (ttl || dpt) {
      user.organizations = [ { title: ttl, department: dpt } ];
    }
  }
  //setPropListValues(user,propName,keyName,key,valueName,values,purgeEmpty);
  setPropListValues(user,'relations','type', 'manager',  'value',info[3], true);
  setPropListValues(user,'relations','type', 'assistant',  'value',info.slice(4,5+1), true);
  setPropListValues(user,'phones', 'type', 'work', 'value',info[6], true);
  setPropListValues(user,'phones', 'type', 'mobile', 'value', info[7], true);
  setPropListValues(user,'ims', 'protocol', 'skype', 'im', info[8],true);
  setPropListValues(user,'ims', 'customProtocol', 'line', 'im', info[9], true);

  if (user.notes) {
    user.notes.value = info[10];
  } else if (info[10]) {
    user.notes = { contentType: 'text_plain', value: info[10] };
  }
  user.includeInGlobalAddressList = Boolean(info[11]);
  user.orgUnitPath = info[12];
  AdminDirectory.Users.update(user, userKey);
}

function dirUpload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DirectorySheetName);
  if (sheet == null) return;
  var isModified = sheet.getRange(2/*row*/,1/*col*/,sheet.getLastRow()-1).getValues();  
  for (var i=0; i<isModified.length; i++) {
    if ('*' == String(isModified[i]).trim()) {
       var userKey = sheet.getRange(i+2, 3).getValue();
       var serializedUserInfo = sheet.getRange(i+2,5,1,headers.length).getValues()[0]; 
       updateUser(userKey, serializedUserInfo);
    }
  }
  // Clear '*' in first column
  // 清除第一欄的 '*'
  sheet.getRange(2,1,sheet.getLastRow()-1).setValue('');
}

function dirDownload() {
  var userInfo = [];
  var options = { }
  var gsDomain = PropertiesService.getDocumentProperties().getProperty('SimpleDirectory.domain');
  if ( (!gsDomain)||(gsDomain=='') ) {
    options.customer = 'my_customer';     // All accounts         所有帳戶
  } else {
    options.domain = gsDomain;            // Specific domain      特定網域
  }

  timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  try {
    do {
      var response = AdminDirectory.Users.list(options);
      response.users.forEach( function(user) {
        userInfo.push(serializeUserInfo(user));
      });
      
      // For domains with many users, the results are paged
      // 如果網域有很多的使用者，回傳資料會分頁
      if (response.nextPageToken) {
        options.pageToken = response.nextPageToken;
      }
    } while (response.nextPageToken);

    userInfo.sort( function(a, b) {
      if (a[orderBy1] == b[orderBy1]) {
        return (a[orderBy2]>b[orderBy2])?1:-1;
      } else {
        return (a[orderBy1]>b[orderBy1])?1:-1;
      } 
    });
  
    getSheet().getRange(2/*row*/,2/*col*/, userInfo.length, userInfo[0].length).setValues(userInfo);
    
  } catch(e) {
    var ui = SpreadsheetApp.getUi();    
    var response = ui.alert(AppTitle, MsgDownloadError.replace('$Domain$', gsDomain), ui.ButtonSet.OK); 
  }
  
}
