# gas-insert-routine-schedules-ts-v2

## 背景
- [gas-insert-routine-schedules](https://github.com/steakyukke/gas-insert-gcal-schedules)ではwebpackでJavaScriptに変換していたが、claspがその役割を担えるようになったので、それにあわせて修正しました。
- さらに、スプレッドシートから実行するのが面倒になったので、Webアプリとして公開し、クエリパラメーターで対象日付を指定できるようにしました。
  - 日付は「今日から●日後」という感じで指定する。たとえば、`[URL]?add=1`なら、明日となる。

## 使い方
[TBD]

## メモ
-  🔗[公開用スプレッドシート](https://docs.google.com/spreadsheets/d/1oMmu-fvZKE3d0zoVQx_MNwKd1bTfFra7_22RTwoTD4Y/edit#gid=1053184431)を利用。


### claspの問題
- `export const`したものが他のファイルから`import`できない。
  - typescript上で`import`できるが、JavaScriptに変換されると`import`行がコメントアウトされる為。
  - `export const xxx=`が`exports.xxx=`と変換されるので、`import`どころか同じファイル内でも利用できなくなってしまう。

### SpreadsheetApp.getActiveSheet()の仕様について
- スプレッドシートにてマクロ実行すると、開いているシートが対象になる。
- スクリプトエディターにて実行すると、スプレッドシートの一番左に位置するシートが対象になる。
- 