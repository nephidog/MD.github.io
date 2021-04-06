# ♚尼腓工作室♚ (一) 利用Google Apps Script(GAS) 整理雲端硬碟資料到Excel

## 前言
我想建一個看動漫的網站，而我的影片是放在雲端硬碟中，雖然可以逐一分享連結來建立網頁但畢竟學過一點程式，還是用聰明一點的方式吧，這邊要先感謝slash跟我說有Google Apps Script(GAS)這個程式，可以利用腳本來完成有規律的命令。
![](https://i.imgur.com/p4YDCbl.png)
Google Apps Script(GAS)能夠連動你的Google帳號，可以輕易的控管雲端硬碟中的檔案，而且執行過程中式使用Google的電腦來運行，不會耗費到自己電腦的效能，真的是太方便拉~
![](https://i.imgur.com/NlGqp5G.png)

## 第一步 創立Google Excel 並開啟 GAS

我希望將整理好的目錄內容放置在Google Excel中，所以我們先到雲端硬碟中創立一個Excel表格，然後再上方的工具點選指令碼編輯器，就會跳轉到我們的腳本介面拉~

![](https://i.imgur.com/RAS6pDK.png)

這就是我們的主角Google Apps Script(GAS)，這邊就像寫JAVA一樣需要寫程式囉~

![](https://i.imgur.com/fHl8BsS.png)

## 第二步 尋找指令
我們可以在[Google Apps Script說明文件](https://developers.google.com/apps-script/reference)中看到支援的程式，並在裡面找到範例程式碼。
![](https://i.imgur.com/lKVM1Dd.png)

點選需要的程式後就可以看到他的指令，可以看到回傳的型態及指令的描述
![](https://i.imgur.com/b7zc7A9.png)

當我們看到可能需要的指令時點選就可以看到範例文建

![](https://i.imgur.com/VYtKZy6.png)

## 第三步 找到目標檔案

在Google Drive中的每一個檔案都有自己專屬的金鑰，我想整理特定資料夾內的影片並輸出到剛剛建立的Excel中，所以我要先找到影片資料夾的金鑰和Excel兩個的金鑰。

我整理的資料如下，(我利用雲端硬碟的自動同步這邊就以電腦端來擷圖)

主資料夾中有各個動漫的資料夾。
![](https://i.imgur.com/SD5g1ef.png)

主資料夾內有所有動漫的資料夾，而動漫資料夾內有自己的影片。
![](https://i.imgur.com/4qfZVFB.png)

我們先找到雲端硬碟，並點到需存放位置，接下來就能在網址列的後面看到一串文字，那就是這個資料夾的金鑰。

![](https://i.imgur.com/36QQVfq.png)

以這邊來說這個資料夾的金鑰就是 `1YlkZ_iBFmUCbCTC_MQGb8DdpYE0bn0hd`

接下來回到Excel找到他的金鑰，一樣先打開Googe Excel但這比較特別金鑰藏在網址中間。

![](https://i.imgur.com/eYrOC79.png)

以這邊來說金鑰就是 `1A_WP-u8x7ywva9SqcEayaPWFuOQx2B1YBCFdeqId5yg`
必要的東西都找到後就可以回到我們的 Google Apps Script(GAS)編寫程式囉。

## 第四步 輸出資料夾名和各動漫名稱
### 測試金鑰輸出
首先宣告置跟Excel金鑰，在這邊宣告是使用var，我們要將指令輸入在他指令碼中
```javascript
function myFunction() {

}
```

首先宣告app為Excel並打上Excel的ID，我希望第一個表格列出所有動漫的資料夾所以宣告sheet為第一個表格，第二章表格列出所以所有的動畫資訊所以宣告sheet2為第二個表格。

```javascript
var app = SpreadsheetApp.openById("希望輸出的Excel金鑰");
var sheet = app.getSheets()[0];//EXCEL第一表格
var sheet2 = app.getSheets()[1];//EXCEL第二表格
```


之後宣告影片資料夾的部分。

```javascript
var videoFolder = DriveApp.getFolderById("資料夾金鑰");//影片資料夾金鑰
var foldersInVideoFolder = videoFolder.getFolders();//獲取目錄中所資料夾的集合
var folder;
var folderID =[];//影片資料夾金鑰
```
利用迴圈在folderID[ ]矩陣中存入影片資料夾的金鑰

``` javascript=1
for (var i = 0; foldersInVideoFolder.hasNext(); i++)  //hasNext()，返回一個項目
    {
        folder = foldersInVideoFolder.next();//next()獲取文件或文件夾集合中的下一項。
        folderID[i] = folder.getId();//在 folderID[]中存影片資料夾金鑰
        var data =[folder.getName(),folder.getId()]
        sheet.appendRow(data);
    }
```
我們利用迴圈來爬取資料，並宣告一個folderID[i]陣列來存放Videod中各影片資料夾中的金鑰，並用data來讀取檔案名稱、及金鑰，而sheet是我們剛剛宣告的第一個表格，appendRow(data)就是新增一行指令，輸出就是data這個變數(資料夾名稱、該資料夾的金鑰)

![](https://i.imgur.com/JVl2o0E.png)

接下來就進行測試，按下專案執行，我們能在Excel中看看輸出的結果拉~!

![](https://i.imgur.com/UVLy3H1.png)

我們成功的利用folderID[　]中的金鑰輸出了資料夾名稱及該資料夾的金鑰，那就可以在進行下一步的改寫，我們先將第5行和第6行註記起來。


```=1
for (var i = 0; foldersInVideoFolder.hasNext(); i++)  //hasNext()，返回一個項目
    {
        folder = foldersInVideoFolder.next();//next()獲取文件或文件夾集合中的下一項
        folderID[i] = folder.getId();//在 folderID[]中存影片資料夾金鑰
        //ar data =[folder.getName(),folder.getId()]
        //sheet.appendRow(data); 
        }
```
### 判斷檔案名稱並輸出

我希望可以判斷檔案名稱判斷他是否式為影片或圖片，這邊利用split來進行切割字串
```
split(".");//分割字串，進行切割以"."來進行切割
```
然後自己製作影片格式與圖片格式的陣列，我只有列出目前有可能有的格式，會用陣列來製作式方便後續如果有新的格式方便管理
```
var ImageFormat=["jpg","jpeg","png","tif","tiff","bmp"]//圖片格式
var Videoformat=["mp4","rmvb"]
```
### 流程圖(單純輸出資料)
![](https://i.imgur.com/yvG2DcH.png)


### 程式碼(單傳輸出資料)
```javascript=1
function myFunction() {

var app = SpreadsheetApp.openById("試算表金鑰");// 想輸出的試算表金鑰
var sheet = app.getSheets()[0];//EXCEL第一表格
var sheet2 = app.getSheets()[1];//EXCEL第二表格

    var videoFolder = DriveApp.getFolderById("資料夾金鑰");//放影片資料夾的資料夾金鑰
    var foldersInVideoFolder = videoFolder.getFolders();//獲取目錄中所資料夾的集合
    var folder;
    var folderID =[];//影片資料夾金鑰

    var ImageFormat=["jpg","jpeg","png","tif","tiff","bmp"]//圖片格式
    var Videoformat=["mp4","rmvb"]

    for (var i = 0; foldersInVideoFolder.hasNext(); i++)
    {
        folder = foldersInVideoFolder.next();//next()獲取文件或文件夾集合中的下一項。
        folderID[i] = folder.getId();//在 folderID[]中存影片資料夾金鑰
        //var data =[folder.getName(),folder.getId()]
        // sheet.appendRow(data);
    }
      for (var i = 0; i < folderID.length; i++){
      var folder = DriveApp.getFolderById(folderID[i]);//folder讀取folderID[]中存影片資料夾金鑰
      var files = folder.getFiles();   //getFiles()獲取目錄中所有文件的集合
      var file; 
      var DataQuantity=0,run=0;
   for(var j=0;files.hasNext();j++){
      file=files.next();
      DataQuantity++;
      dotSplit = file.getName().split("."); //dotSplit取得文件名稱，並以"."來分割
      

    var count=0; count2=0//count圖片格式，count2影片格式
    
    while((count<ImageFormat.length)){
    
    if(dotSplit[dotSplit.length-1]==ImageFormat[count]){//判斷副檔名是否為影片格式
    
    var data =[folder.getName(),folder.getId(),file.getName(),file.getUrl()]
    sheet.appendRow(data);
    count=ImageFormat.length;//成功輸出過就暫停
    
    }else{count++,run++
    
    while(count2<Videoformat.length){
    
    if(dotSplit[dotSplit.length-1]==Videoformat[count2]){//判斷是否為影片
    var data =[folder.getName(),file.getId(),file.getName(),"https://drive.google.com/file/d/"+file.getId()+"/preview"]
    sheet2.appendRow(data);
    count2=Videoformat.length;
    }else{count2++;}
    }
    }   //繼續判斷副檔名
    }
}
if(DataQuantity==run/ImageFormat.length){//如果資料夾中沒有圖片，輸出資料夾名稱跟金鑰
    var data =[folder.getName(),folder.getId()]
    sheet.appendRow(data);
    }
    }
}
```
## 第五步 更新排除重複資料

按照第四部已經可以照我的要求跑出資料了，但如果我有放入新的影片，這時我重新啟動腳本，就會發現以前輸出過的資料又輸出了一次，所以要加入判斷，避免重複資料。
### 會用到的函式

這邊會利用到將字串分割來進行判斷檔案類型，以及將試算表中的目錄名稱及影片檔案名稱記錄在陣列中方便後續比對，最後將輸出完的進行排序。

```javascript
getDataRange();//取得取得表格範圍
```

```javascript
getValues();//取得內容中的資料
```
```javascript
sort();//排序資料
```
### 流程圖(不重複資料)
![](https://i.imgur.com/i5EDwGj.png)
### 程式碼(不重複資料)
```javascript=1
function myFunction() {

var app = SpreadsheetApp.openById("試算表金鑰");// 試算表金鑰
app.setName("動漫資料庫")
var sheet = app.getSheets()[0];//EXCEL第一表格(資料夾)[資料夾名、資料夾金鑰、圖片名、圖片網址]
var range = sheet.getDataRange();//取得表格內容
var values = range.getValues();//取得內容中的資料
sheet.setName("動漫資料夾")
var sheet2 = app.getSheets()[1];//EXCEL第二表格(影片總級數)[資料夾名、、影片金鑰、影片名、影片網址]
var range2 = sheet2.getDataRange();//取得表格內容
var values2 = range2.getValues();//取得內容中的資料
sheet2.setName("影片資料庫")

var ImageFormat=["jpg","jpeg","png","tif","tiff","bmp"]//圖片格式
var Videoformat=["mp4","rmvb"]//影片格式

//存取表單現有內容

var FolderName=[];
for (var i = 0; i < values.length; i++) {
  FolderName[i] = values[i][0];
}
var VideoName=[];
for (var i = 0; i < values2.length; i++) {
  VideoName[i] = values2[i][2];
}

    var videoFolder = DriveApp.getFolderById("影片資料夾金鑰");//影片資料夾金鑰
    var foldersInVideoFolder = videoFolder.getFolders();//獲取目錄中所資料夾的集合
    var folder;
    var folderID =[];//影片資料夾金鑰

    for (var i = 0; foldersInVideoFolder.hasNext(); i++){
        folder = foldersInVideoFolder.next();//next()獲取文件或文件夾集合中的下一項。
        folderID[i] = folder.getId();//在 folderID[]中存影片資料夾金鑰
        //var data =[folder.getName(),folder.getId()]
         //sheet.appendRow(data);
    }
      for (var i = 0; i < folderID.length; i++){
      var folder = DriveApp.getFolderById(folderID[i]);//folder讀取folderID[]中存影片資料夾金鑰
      var files = folder.getFiles();//getFiles()獲取目錄中所有文件的集合
      var file; 
      var DataQuantity=0,run=0;
   for(var j=0;files.hasNext();j++){
      file=files.next();
      DataQuantity++;
      dotSplit = file.getName().split("."); //dotSplit取得文件名稱，並以"."來分割
      
    
    var count=0; count2=0//count圖片格式，count2影片格式
    
    while((count<ImageFormat.length)){
    
    if(dotSplit[dotSplit.length-1]==ImageFormat[count]){//判斷副檔名是否為圖片格式
    var iia =0;
    while(iia<FolderName.length){
    if(folder.getName()==FolderName[iia]){
    sheet.getRange(iia+1,1).setValue(folder.getName());
    sheet.getRange(iia+1,2).setValue(folder.getId());
    sheet.getRange(iia+1,3).setValue(file.getName());
    sheet.getRange(iia+1,4).setValue("https://drive.google.com/file/d/"+file.getId()+"/preview");
    iia=FolderName.length;
    }else{//沒有相同名稱
    iia++
    if(iia==FolderName.length){
    var data =[folder.getName(),folder.getId(),file.getName(),file.getUrl()]
    sheet.appendRow(data);
    }
    }
    }
    //var data =[folder.getName(),folder.getId(),file.getName(),file.getUrl()]
    //sheet.appendRow(data);
    count=ImageFormat.length;//成功輸出過就暫停
    
    }else{count++,run++
    
    while(count2<Videoformat.length){
    
    if(dotSplit[dotSplit.length-1]==Videoformat[count2]){//判斷是否為影片
    var iic=0;
    while(iic<VideoName.length){
    if(file.getName()==VideoName[iic]){
    sheet2.getRange(iic+1,1).setValue(folder.getName());
    sheet2.getRange(iic+1,2).setValue(file.getId());
    sheet2.getRange(iic+1,3).setValue(file.getName());
    sheet2.getRange(iic+1,4).setValue(file.getUrl());
    iic=VideoName.length;
    }else{
    iic++;
    if(iic==VideoName.length){
    var data =[folder.getName(),file.getId(),file.getName(),file.getUrl()]
    sheet2.appendRow(data);
    }
    }
    }
    //var data =[folder.getName(),file.getId(),file.getName(),file.getUrl()]
    //sheet2.appendRow(data);
    count2=Videoformat.length;
    }else{count2++;}
    }
    }   //繼續判斷副檔名
    }  
}

if(DataQuantity==run/ImageFormat.length){//如果資料夾中沒有圖片，判斷是否重複
  var iib=0;
  while(iib<FolderName.length){
  if(folder.getName()==FolderName[iib]){
sheet.getRange(iib+1,1).setValue(folder.getName());
sheet.getRange(iib+1,2).setValue(folder.getId());

  iib=FolderName.length;
  }else{//沒有相同名稱就新增
  iib++;
  if(iib==FolderName.length) {
  var dara=[folder.getName(),folder.getId()];
  sheet.appendRow(dara);}
  }    
    } 
      }    
        }
    //因為影片很亂所以來做排序
    sheet.sort(1);//影片資料夾名按照資料夾名稱排序
    sheet2.sort(3).sort(1);//動漫資料庫先以片名排序後再依照資料夾名稱排序
}
```
執行專案後，可以發現，在Excel中並沒有重複的資料，而且有按照檔案名稱來升序排序

影片目錄
![](https://i.imgur.com/Q4KKpK7.png)

影片資料庫
![](https://i.imgur.com/GhQj6DA.png)

目前就有4XX多筆資料如果這些是手動整理真的會累死~有腳本真是方便
## 第六步 自動更新及效能
已經可以做到重覆執行就可以更新內容，接下來就可以讓他自動更新，我們可以到[專案管理](https://script.google.com/home)的地方右鍵點選觸發條件
![](https://i.imgur.com/zIwpplP.png)

選擇右下角的新增觸發條件
![](https://i.imgur.com/J1D34ma.png)
![](https://i.imgur.com/8uV8kVR.png)
可以設定特定時間就執行一次，我如果需要更新只需要將影片放入資料夾中，腳本到規定時間就會自動幫我整理囉~

我們可以在[我的專案中](https://script.google.com)，點選我的執行項目
![](https://i.imgur.com/LFx1Sla.png)

在我的執行項目中能看到腳本執行的時間，以及是否執行完成
![](https://i.imgur.com/3qeX0mk.png)

向這邊我的腳本在第一次執行時花了165秒左右，第二次之後就變成75秒多，如果不小心寫成無限迴圈就只能等他自己停止了或者砍掉試算表重新編寫，所以邊寫腳本時要小心一點。
## 結語
要把自己想到的事情轉換成程式語言真的是一個困難的事情，這邊有許多的邏輯判斷花了我好多時間才完成，看來我的程式能力還是要在多多練習，到這裡動漫資料庫的整理已經完成了，下一步就是建立資料庫並將資料塞進去，我們下次見拉~
![](https://i.imgur.com/OiloIhe.png)
