# excel002

軽量な **Excel（.xlsx 等）の読み込み・書き込み** を目的とした Java **ライブラリ**です。**Apache POI および FastExcel は使用しません。** 標準ライブラリ中心でストリーミング読み書きできる構成を目指します。

大きなブックでは、ワークブック全体をメモリに載せず、OOXML（ZIP 内のシート XML など）を逐次処理してヒープ使用量を抑えます。

## 要件

- JDK 17 以上（ビルドは `release 17` でコンパイルします）
- Apache Maven 3.x

## ビルドとインストール

通常の JAR（依存は別途）として `target/excel002-1.0-SNAPSHOT.jar` が出力されます。

```bash
mvn package
```

ローカルの Maven リポジトリに入れて他プロジェクトから参照する場合:

```bash
mvn install
```

## 利用側プロジェクトでの読み込み（Maven）

```xml
<dependency>
    <groupId>jp.engawa</groupId>
    <artifactId>excel002</artifactId>
    <version>1.0-SNAPSHOT</version>
</dependency>
```

`mvn install` 済みであるか、`repositories` でこのプロジェクトが公開されているリポジトリを指定してください。

## クラスパスでの利用

ビルドした JAR をクラスパスに追加します。

```bash
java -cp target/excel002-1.0-SNAPSHOT.jar jp.engawa.excel002.App
```

（`App` は動作確認用のエントリです。ライブラリ利用時は自アプリケーションのコードから API を呼び出してください。）

## `excel002.properties`

`ExternalConfig.load(Class<?> anchor)` で読み込みます。**`anchor` には利用側アプリケーションのクラス**（実行環境で読み込まれている JAR の場所を決めるため）を渡してください。ファイルが無い場合は警告を出しつつ空のプロパティで続行します。

テスト用サンプルとしてリポジトリ直下に `excel002.properties` を置いています。
