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
 
/*

var DirectorySheetName = 'Users';         // Sheet name
var headers = ['Full name',
               'Primary e-mail',
               'Last log-in',
               'Employee Id',
               'Title',
               'Department',
               'Manager',
               'Assistant 1',
               'Assistant 2',
               'Work phone',
               'Mobile phone',
               'Skype',
               'Line',
               'Note',
               'Global',
               'Org path'];

var orderBy1 = 15; // 1st sort oder = "org path"
var orderBy2 = 0;  // 2nd sort order = "full name"

var AppTitle = 'SimpleDirectory';                                          // Title in dialog
var MsgInitDomain = 'Please enter G Suite domain (currently "$Domain$")';  // Message in "Set up domain" dialog
var MsgDownloadError = 'Cannot access "$Domain$" directory';               // Error message
var MsgNeverLoggedIn = 'Never logged in';

var MenuDownload = '⭳ Download directory';
var MenuUpload = '⬆ Update directory';
var MenuSetup = '⚙ Set up';

*/

var DirectorySheetName = 'Users'; // 工作表名稱

var headers = ['全名','主要電郵','最後一次登入時間','員工編號','職稱','部門','主管','助理 1','助理 2','公司電話','行動電話','Skype','Line', '附註','公開','組織路徑'];

var orderBy1 = 15; // 依照'組織路徑'作主要排序
var orderBy2 = 0;  // 依照'全名'作主要排序

var AppTitle = 'SimpleDirectory';                            // 對話窗的標題
var MsgInitDomain = '請輸入 G Suite 網域 (目前是 "$Domain$")';  // 「設定」對話窗的訊息
var MsgDownloadError = '無法存取 "$Domain$" 目錄';             // 錯誤訊息
var MsgNeverLoggedIn = '從未登入';

var MenuDownload = '⭳ 載入目錄';
var MenuUpload = '⬆ 上傳資料';
var MenuSetup = '⚙ 設定';
