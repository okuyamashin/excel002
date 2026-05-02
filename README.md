# excel002

軽量な **Excel（.xlsx 等）の読み込み・書き込み** を目的とした Java プロジェクトです。依存を抑えつつ、CLI やツールとして動かせる構成を想定しています。

## 要件

- JDK 17 以上（ビルドは `release 17` でコンパイルします）
- Apache Maven 3.x

## ビルド

```bash
mvn package
```

実行可能な fat JAR は `target/excel002-1.0-SNAPSHOT.jar` に生成されます。

## 実行

fat JAR と**同じディレクトリ**に `excel002.properties` を置くと、アプリ起動時に読み込みます（IDE やクラスパス直下から実行する場合は、カレントディレクトリのファイルを参照します）。

```bash
java -jar target/excel002-1.0-SNAPSHOT.jar
```

設定ファイルが無い場合は警告を出しつつ空の設定で続行します。
