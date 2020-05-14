# ZoomURLGetter-for-TMU
## 概要
Kibacoからtmuメールアドレスに送信された、zoomのリンクを取得、一覧表示します。  
Kibacoが落ちててZoomURLが分からない！という非常事態のときに使えます(たぶん)
## ダウンロード
https://1drv.ms/u/s!ArNbpC7I1GPWgp4SONrZPPXz8auEUQ?e=zqIMbt

## 仕様
・Windowsのみ対応です。
・現在のところ一覧表示機能のみです。  
・Kibacoから自動送信されたメールのみ対応しています。　　  
・5/1~のメールを表示します。  
## 使用方法
1.フォルダ内のZoomURLGetter.exeを起動  
2.右上の"Show Mail List"ボタンをクリック  
3.tmuのアドレス(末尾が@ed.tmu.ac.jp)のアドレスでログインします。  
4.Resultボックスにメール一覧が表示されます。  
5.メールを表示ボタンで新ウィンドウでメール内容が表示されます。  

### singn-outボタン
sign-outボタンを押すとサインアウトします。  
通常、sing-outボタンを押さない限り次回起動時もログイン状態は保持されます。

##  開発環境
・Visual Studio 2019 Community  
・.Net Frame Work 4.7.2  
・Microsoft Graph API v 1.0  
・Microsoft Graph .NET Client Library  

##  参考
[MicrosoftGraph](https://docs.microsoft.com/ja-jp/graph/overview)  
[Microsoft Graph .NET Client Libraryを使ったAPIの呼び出しサンプル](https://www.ka-net.org/blog/?p=10169)  

