![Hidennotare](https://raw.githubusercontent.com/wiki/RelaxTools/Hidennotare/image/Hidennotare.png)

### Secret Source VBA Library for Excel 2010/2013/2016/2019/365 32/64bit

### 目次
#### 1. 使用方法
[1.1. ICursor 系コマンドの使い方](https://github.com/RelaxTools/Hidennotare/wiki/1.1.-ICursor-%E7%B3%BB%E3%82%B3%E3%83%9E%E3%83%B3%E3%83%89%E3%81%AE%E4%BD%BF%E3%81%84%E6%96%B9)  
[1.2. StringBuilderの使い方](https://github.com/RelaxTools/Hidennotare/wiki/1.2.-StringBuilder-%E3%81%AE%E4%BD%BF%E3%81%84%E6%96%B9)  
[1.3. Messageクラスの使い方](https://github.com/RelaxTools/Hidennotare/wiki/1.3.-Message%E3%82%AF%E3%83%A9%E3%82%B9%E3%81%AE%E4%BD%BF%E3%81%84%E6%96%B9)  
[1.4. Usingクラスの使い方](https://github.com/RelaxTools/Hidennotare/wiki/1.4.-Using%E3%82%AF%E3%83%A9%E3%82%B9%E3%81%AE%E4%BD%BF%E3%81%84%E6%96%B9)  
[1.5. CSV Parser の使い方](https://github.com/RelaxTools/Hidennotare/wiki/1.5.-CSV-Parser-%E3%81%AE%E4%BD%BF%E3%81%84%E6%96%B9)  
[1.6. ArrayListとDictionaryのシリアライズ化](https://github.com/RelaxTools/Hidennotare/wiki/1.6.-ArrayList%E3%81%A8Dictionary%E3%81%AE%E3%82%B7%E3%83%AA%E3%82%A2%E3%83%A9%E3%82%A4%E3%82%BA%E5%8C%96)  
[1.7. ArrayListとDictionaryのデシリアライズ化](https://github.com/RelaxTools/Hidennotare/wiki/1.7.-ArrayList%E3%81%A8Dictionary%E3%81%AE%E3%83%87%E3%82%B7%E3%83%AA%E3%82%A2%E3%83%A9%E3%82%A4%E3%82%BA%E5%8C%96)  
#### 2. リファレンス
##### 2.1. インターフェイス
[2.1.1. IConstructor インターフェース](https://github.com/RelaxTools/Hidennotare/wiki/2.1.1.-IConstructor%E3%82%A4%E3%83%B3%E3%82%BF%E3%83%BC%E3%83%95%E3%82%A7%E3%82%A4%E3%82%B9)  
[2.1.2. ICursor インターフェース](https://github.com/RelaxTools/Hidennotare/wiki/2.1.2.-ICursor-%E3%82%A4%E3%83%B3%E3%82%BF%E3%83%BC%E3%83%95%E3%82%A7%E3%82%A4%E3%82%B9)  
[2.1.3. IList インターフェース](https://github.com/RelaxTools/Hidennotare/wiki/2.1.3.-IList-%E3%82%A4%E3%83%B3%E3%82%BF%E3%83%BC%E3%83%95%E3%82%A7%E3%83%BC%E3%82%B9)  
[2.1.4. IDictionary インターフェース](
https://github.com/RelaxTools/Hidennotare/wiki/2.1.4.-IDictionary-%E3%82%A4%E3%83%B3%E3%82%BF%E3%83%BC%E3%83%95%E3%82%A7%E3%82%A4%E3%82%B9)  
##### 2.2. クラス
[2.2.1. ArrayList クラス](https://github.com/RelaxTools/Hidennotare/wiki/2.2.1.-ArrayList-%E3%82%AF%E3%83%A9%E3%82%B9)  
[2.2.2. Dictionary クラス](https://github.com/RelaxTools/Hidennotare/wiki/2.2.2.-Dictionary-%E3%82%AF%E3%83%A9%E3%82%B9)  
[2.2.3. OrderedDictionary クラス](https://github.com/RelaxTools/Hidennotare/wiki/2.2.3.-OrderedDictionary-%E3%82%AF%E3%83%A9%E3%82%B9)  
[2.2.4. SortedDictionary クラス](https://github.com/RelaxTools/Hidennotare/wiki/2.2.4.-SortedDictionary-%E3%82%AF%E3%83%A9%E3%82%B9)  

#### 3. 機能一覧

|No.|クラス|機能|
|---|---|---|
|1|Core.bas|唯一の標準モジュール|
|2|ArrayList.cls|ArrayListクラス(拡張Collection)|
|3|BookReader.cls|ブック読み込みクラス|
|4|BookWriter.cls|ブック書き込みクラス|
|5|Callback.cls|OnAction/OnTime等コールバッククラス|
|6|CharCursor.cls|文字単位カーソルクラス|
|7|CheckDigit.cls|チェックディジットクラス|
|8|Clipboard.cls|クリップボードクラス|
|9|CollectionCursor.cls|コレクションカーソルクラス|
|10|Convert.cls|変換クラス|
|11|CSVReader.cls|CSV Readerクラス|
|12|CsvWriter.cls|CSV Writerクラス|
|13|Dialog.cls|ダイアログクラス|
|14|Dictionary.cls|拡張Dictionaryクラス|
|15|ExcelHelper.cls|Excel固有の処理をまとめたクラス|
|16|ExplorerComparer.cls|Explorerと同様の並び替えクラス|
|17|FileIO.cls|ファイルIOクラス(FileSystemObject拡張）|
|18|FormManager.cls|フォーム実行中・キャンセル処理クラス|
|19|IComparer.cls|比較用インターフェース|
|20|IConstructor.cls|コンストラクタインターフェース|
|21|ICursor.cls|カーソルインターフェース|
|22|IDictionary.cls|Dictionaryインターフェース|
|23|IDictionaryAccessor.cls|Dictionaryアクセサインターフェース|
|24|IIterator.cls|イテレータインタフェース|
|25|IList.cls|ArrayList インターフェース|
|26|INewInstance.cls|インスタンス生成インターフェース|
|27|IReader.cls|Readerインターフェース|
|28|IUsing.cls|Using インターフェース|
|29|IWriter.cls|Writerインターフェース|
|30|JSON.cls|JSON パースクラス|
|31|LineCursor.cls|文字列配列のカーソルクラス|
|32|Logger.cls|ログ出力クラス|
|33|MatchBean.cls|通常検索・正規表現結果クラス|
|34|Math.cls|計算クラス|
|35|MCommand.cls|PowerQuery M 言語Command 生成クラス|
|36|MCsv.cls|PowerQuery M 言語CSV 生成クラス|
|37|Message.cls|メッセージ出力クラス|
|38|MFile.cls|PowerQuery M 言語File 生成クラス|
|39|MList.cls|PowerQuery M 言語List 生成クラス|
|40|MRecord.cls|PowerQuery M 言語Record 生成クラス|
|41|MTable.cls|PowerQuery M 言語Table 生成クラス|
|42|NewExcel.cls|マルチプロセス起動クラス(Excel)|
|43|NewPowerPoint.cls|マルチプロセス起動クラス(PowerPoint)|
|44|NewWord.cls|マルチプロセス起動クラス(Word)|
|45|OneTimeSpeedBooster.cls|スピードブースタークラス|
|46|OrderedDictionary.cls|順序保証Dictionaryクラス|
|47|PairLogger.cls|前後ロガークラス(Trace)|
|48|PairLoggerInfo.cls|前後ロガークラス(Info)|
|49|Process.cls|プロセス起動|
|50|RangeCursor.cls|Rangeカーソルクラス|
|51|RangeHelper.cls|Rangeヘルパークラス|
|52|RegExp.cls|正規表現クラス|
|53|Registry.cls|レジストリ操作クラス|
|54|SheetCursor.cls|シート内容カーソルクラス|
|55|SheetIterator.cls|シート内容イテレータクラス|
|56|SortedDictionary.cls|ソート対応Dictionaryクラス|
|57|Stack.cls|スタッククラス|
|58|StdSch.cls|標準検索クラス|
|59|StringBuilder.cls|StringBuilderクラス|
|60|StringHelper.cls|文字列操作クラス|
|61|TableCursor.cls|テーブルカーソルクラス|
|62|TaskTrayView.cls|タスクトレー表示クラス|
|63|TextReader.cls|テキストReaderクラス|
|64|TextWriter.cls|テキストWriterクラス|
|65|Using.cls|Using クラス|
|66|Web.cls|Web関連クラス|
|67|frmShowOnece.frm|次回から表示しないフォーム|
|68|frmWait.frm|実行中フォーム|
