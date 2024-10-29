# vba-calculate-man-hours


### ＜前提条件＞

UTF8toANSI.batを使用し、memo_yyyymmdd.txtをUTF-8ではなく、  
ANSI保存した状態で実行すること  
※文字化けして文字列処理がマッチングしない  

工数算出_step1_抽出.basを「Sheet1」シートにて実行する  
工数算出_step2_入力.basを「yyyymm」（※例 202410）シートにて実行する  

### ＜工数メモ＞

memo_yyyymmdd.txtのサンプル
```.txt
2024年10月10日　10:43

s

現場出勤



2024年10月10日　11:00

会議対応


2024年10月10日　12:00


昼


2024年10月10日　13:00

```