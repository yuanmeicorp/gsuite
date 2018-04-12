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
 
function getEmployId(user) {
  return getPropListValues(user, 'externalIds', 'type', 'organization', 'value')[0];
}

function getManager(user) {
  return getPropListValues(user, 'relations', 'type', 'manager', 'value')[0];
}

function getAssistants(user, maxResult) {
  return getPropListValues(user, 'relations', 'type','assistant', 'value', maxResult);
}

function getPhones(user, type, maxResult) {
  return getPropListValues(user, 'phones', 'type', type, 'value', maxResult);
}

function getIMs(user, protocol, maxResult) {
  var defList = ["aim", "gtalk", "icq", "jabber", "msn", "net_meeting", "qq", "skype", "yahoo"];
  if (defList.indexOf(protocol) != -1) {
    return getPropListValues(user, 'ims', 'protocol', protocol, 'im', maxResult);
  } else {
    return getPropListValues(user, 'ims', 'customProtocol', protocol, 'im', maxResult);
  }
}
              

/**
 * 產生預先填好的陣列
 * @param {Object} value  數值
 * @param {number} length 陣列長度
 * @returns {Array}
 */
function makeArrayOf(value, length) {
  // ES6 ** App Script 目前不支援 **
  // return Array(length).fill(value)
  
  // ES3
  var arr = [], i = length;
  while (i--) { arr[i] = value; }
  return arr;
}

/**
 * 更新 user 的屬性陣列內符合收尋條件的項目的數值
 * @param {Object} user 目錄人員
 * @param {String} propName 屬性陣列名稱
 * @param {String} keyName 索引名稱
 * @param {String} key 索引值
 * @param {String} valueName 值名稱
 * @param {String|Array} values 值/值陣列
 * @param {boolean} [purgeEmpty=false] 是否刪除無數值的屬性
 * @returns {Array}
 */
function setPropListValues(user, propName, keyName, key, valueName, values, purgeEmpty) {

   if (typeof purgeEmpty != 'boolean') { purgeEmpty = false; }

   var props = user[propName];
   var pIdx;   // props index
   var vIdx;   // values index
   var val; // temporary variable
   
   if (values instanceof Array) {
      values = values.map( function(v) { return String(v).trim() } );
   } else {
      values = [String(values).trim()];
   }

   function appendRest(idx){
     for (var k = idx; k < values.length; k++) {
       if (values[k].length) {
         var obj = new Object;
         obj[keyName] = key;
         obj[valueName] = values[k];
         props.push(obj);
       }
     } // for
   } // appendRest
   
   if (props instanceof Array) {
     pIdx = 0;
     vIdx = 0;
     while (pIdx < props.length) {
       var match = (props[pIdx][keyName]) && (props[pIdx][keyName] == key); // keyMatch  索引值吻合
       
       if (match) {
         var purge = (values[vIdx].length) // 空數值
                   && purgeEmpty           // 是否刪除空數值?
                   && (props.length > 1);  // Directory API bug: 如果陣列是空的，屬性不會更新
         if (purge) {
           props.splice(pIdx,1);
         } else {
           props[pIdx++][valueName] = values[vIdx];
         }
         vIdx++;
       } else {
         pIdx++;
       }
       
     } // while
     
   } else {
     vIdx = 0;
     props = user[propName] = [];     
   }
   
   appendRest(vIdx);
}

/**
 * 搜索 user 屬性列表內符合收尋條件的項目的數值
 * @param {Object} user 目錄人員
 * @param {String} propName 屬性陣列名稱
 * @param {String} keyName 索引名稱
 * @param {String} key 索引值
 * @param {String} valueName 值名稱
 * @param {number} [maxResult=1] 最多的結果
 * @returns {Array}
 */
function getPropListValues(user, propName, keyName, key, valueName, maxResult) {

  if ((typeof maxResult != 'number') || (maxResult <= 0)) { maxResult = 1; }
  
  var props = user[propName];
  var values = makeArrayOf('',maxResult);
  
  if (props instanceof Array) {
    var pIdx = 0; // props index
    var vIdx = 0; // values index
    
    for( ; pIdx < props.length ; pIdx++) {
      var match = (props[pIdx][keyName]) && (props[pIdx][keyName] == key); // keyMatch  索引值吻合
      if (match) {
        values[vIdx++] = props[pIdx][valueName]||'';
        if (vIdx >= maxResult) break;
      }
    }
  }
  return values;
}
