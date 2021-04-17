# gas-insert-routine-schedules-ts-v2

## 背景
- [gas-insert-routine-schedules](https://github.com/steakyukke/gas-insert-gcal-schedules)ではwebpackでJavaScriptに変換していたが、claspがその役割を担えるようになったので、それにあわせて修正しました。

## スプレッドシートからの使い方

## メモ
-  🔗[公開用スプレッドシート](https://docs.google.com/spreadsheets/d/1oMmu-fvZKE3d0zoVQx_MNwKd1bTfFra7_22RTwoTD4Y/edit#gid=1053184431)を利用。
- 使い方など→　🔗[gas-insert-routine-schedules](https://github.com/steakyukke/gas-insert-gcal-schedules)


### claspの問題
- `export const`したものが他のファイルから`import`できない。
  - typescript上で`import`できるが、JavaScriptに変換されると`import`行がコメントアウトされる為。
  - `export const xxx=`が`exports.xxx=`と変換されるので、`import`どころか同じファイル内でも利用できなくなってしまう。
