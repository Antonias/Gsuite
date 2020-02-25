var scriptProperties = PropertiesService.getScriptProperties();

/*
     *チャットワーク取り扱いをまとめたクラス
     *2020/2/21作成　現時点ではチャットルームへの書き込みのみを実行
*/   
(function(global){
  var chatwork = (function(){
    
    /**
     *コンストラクタ
     *以下の設定を行う
     *token_guite：Gsuite側のプロパティ用トークン
     *token_chatwork：チャットワーク側トークン
     *spread_sheet_url：タイトル⇔ルームID紐付け用スプレッドシートのURL
     *administrator：スケジュール管理者
     */   
    function chatwork(config){
      this.token_gsuite = config.token_gsuite
      this.token_chatwork = config.token_chatwork      
      this.spread_sheet_url = config.spread_sheet_url
      this.administrator = config.administrator
    };

    /**
     * メイン処理
     */
    chatwork.prototype.main = function(e){
      var calendarId = e.calendarId;
      // 予定取得時にsyncTokenを指定して差分イベントを取得
      var optionalArgs = {
        'syncToken': this.getNextSyncToken(calendarId)
      };
     
      var events = Calendar.Events.list(calendarId, optionalArgs);  
      console.log('取得した予定数:%s', events.items.length);

      for (var i = 0; i < events.items.length; i++) {
        var event = events.items[i];
        console.log('event.summary:%s/event.start:%s/event.end:%s/states:%s　else:%s', event.summary, event.start, event.end, event.sequence, Date.parse(event.created) - Date.parse(event.updated));
        //タイトルが取ってこれた(get_title)＋ルームIDを取得できた(get_room_id)分をチャットワーク側に記載
        title = this.get_title_from_time(event)//イベントは作成、削除、更新とあるため、それぞれで記載するかどうかと、タイトルを設定する
        if(title[0] == true){
          room_id = this.get_chatwork_id(event)//要約によってルームIDと、記載するかどうかを判定する
          if (room_id[0]){
            this.write_info_to_chatroom(title[1],room_id[1],event)                        
          }  
        }
      }
     this.saveNextSyncToken(events.nextSyncToken)
    };
    
    /**
     *チャットルームへの書き込み    
     */   
    chatwork.prototype.write_info_to_chatroom = function(title,room_id,event){
      var client = ChatWorkClient.factory({token: this.token_chatwork});　//チャットワークAPI
        client.sendMessage({
        room_id:room_id, //ルームID
          body: title +'\nタイトル:' + event.summary　+ '\n' +          
          '管理者：' + this.administrator + '\n' +
          '作成者：' + event.creator['email'] +
        '\n日時:' + event.start['dateTime'].slice(0,16).replace('T',' ') + '～' + 
                    event.end['dateTime'].slice(11,16).replace('T',' ')});                 
    };
    
    /**
     *チャットワーク側へ書き込む際のタイトル設定
     *0:書き込み可かどうか 1：書き込む際の要約
     */       
    chatwork.prototype.get_title_from_time = function(event){
      var createDate = event.created
      var updateDate = event.updated
      
      Logger.log('time is ' + createDate + 'and ' + updateDate)
      if(Date.parse(updateDate) - Date.parse(createDate) > 2000){
        return [true,'スケジュールが変更されました.'];}
      else if(Date.parse(updateDate) - Date.parse(createDate) <= 2000){
        return [true,'スケジュールが作成されました.'];}
      else{
        return [false,''];}
    };

    
    /**
     * タイトルからルームIDの設定
     */
    chatwork.prototype.get_chatwork_id = function(event){
      var spreadsheet = SpreadsheetApp.openByUrl(this.spread_sheet_url);
      var sheets = spreadsheet.getSheetByName('マスタ')
      i = 2
      
      do {
        keyword_1 = String(sheets.getRange(i, 1).getValue())
        keyword_2 = String(sheets.getRange(i, 2).getValue())
        id = String(sheets.getRange(i, 3).getValue())
        
        if (event.summary.indexOf(keyword_1)!==-1 & event.summary.indexOf(keyword_2)!==-1){
          return [true, id];   //titleが含まれている場合、書き込みフラグとルームIDを指定
        }
        i = i + 1
      }while (sheets.getRange(i, 1).getValue() != "");
      
      return [true, '175020403']//表示させる場合スケジュール変更のルームに表示
    };
    
    /**
    * 差分トークンの取得
    */
    chatwork.prototype.getNextSyncToken = function(calendarId) {
      // ScriptPropetiesから取得
      var nextSyncToken = scriptProperties.getProperty(this.token_gsuite);
      if (nextSyncToken) {
        console.log('getNextSyncToken(from property):%s', nextSyncToken);
        return nextSyncToken
      }
      
      // ScriptPropetiesにない場合は、カレンダーから取得
      var events = Calendar.Events.list(calendarId, {'timeMin': (new Date()).toISOString()}); 
      nextSyncToken = events.nextSyncToken;
      console.log('getNextSyncToken(from calendar):%s', nextSyncToken);
      return nextSyncToken;
    };
    
    /**
     * 差分トークンの保存
     */
    chatwork.prototype.saveNextSyncToken =function(nextSyncToken) {
      console.log('saveNextSyncToken:%s', nextSyncToken);
      scriptProperties.setProperty(this.token_gsuite, nextSyncToken);
    };
    
    
  return chatwork;
  })();
  
  global.chatwork = chatwork;
})(this);



