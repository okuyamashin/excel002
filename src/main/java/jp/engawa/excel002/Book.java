package jp.engawa.excel002;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URISyntaxException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.StandardCopyOption;
import java.nio.file.attribute.BasicFileAttributes;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.HexFormat;
import java.util.Properties;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

public class Book {
    private static final int STREAM_BUFFER_SIZE = 8192;

    public static Book load(File filepath) throws IOException {
        Book book = new Book();
        book.excelfile = filepath;
        book.beginTime = System.currentTimeMillis();
        book.endTime = 0;
        book.unzip();
        book.sheetsjson();
        book.stringsndjson();
        book.sheetvalues();
        book.sheettypes();
        book.sheetstrings();
        book.sheetformulas();
        book.sheetformats();
        book.sheetmerges();
        book.sheetstyles();
        book.stylesjson();
        book.sheetformattedvalues();
        return book;
    }

    /**
     * {@code excel002.properties} の {@code tmp.dir}（未設定時はプロジェクト慣例どおり {@code tmp}）を絶対パスに解決し、
     * そのディレクトリーが存在すれば**配下を含めて再帰的にすべて削除**する（ルートの {@code tmp.dir} ディレクトリ自体も削除する）。
     *
     * <p>存在しない場合は何もしない。
     */
    public static void clearAllTmpDir() throws IOException {
        Properties props;
        try {
            props = ExternalConfig.load(Book.class);
        } catch (URISyntaxException e) {
            throw new IOException("設定の解決に失敗しました", e);
        }
        Path baseTmp = Path.of(props.getProperty("tmp.dir", "tmp")).toAbsolutePath().normalize();
        deleteRecursively(baseTmp);
    }

    protected File excelfile;
    protected File tmpdir;
    protected File logdir;
    protected File outputdir;
    protected long beginTime;
    protected long endTime;
    protected String status = "none";
    protected String md5;

    public String getStatus() {
        return this.status;
    }

    public long getBeginTime() {
        return this.beginTime;
    }

    public long getEndTime() {
        return this.endTime;
    }

    /**
     * {@code tmp.dir/md5/data} を作成し、その配下に {@code sheets.json} を書き出す。
     * JSON の構造はコード内コメント（実装時仕様）どおり:
     * {@code file_name}, {@code last_modified}（エポックミリ秒）, {@code file_md5},
     * {@code sheets_size}, {@code sheets[]}（各要素に {@code sheet_name}, {@code sheet_index},
     * {@code max_row_num}, {@code max_column_num}, {@code sheet_id}, {@code relationship_id},
     * {@code worksheet_xml} は {@code tmp.dir/md5/excel} からの相対パス（例: {@code xl/worksheets/sheet1.xml}）。JSON のキー名は {@code worksheet_xml} のまま。
     */
    public Book sheetsjson() throws IOException {
        this.status = "sheets.json";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }
        if (excelfile == null || !excelfile.isFile()) {
            throw new IOException("元の Excel ファイルがありません");
        }

        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        Path jsonPath = mdRoot.resolve("data").resolve("sheets.json");
        SheetsJsonExporter.write(excelRoot, jsonPath, excelfile, md5);

        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir/md5/data/strings.ndjson} に {@code xl/sharedStrings.xml} の各 {@code si} を、{@code id}（0 始まりインデックス）と
     * {@code value} の NDJSON で書き出す。
     *
     * @see StringsNdjsonExporter#write
     */
    public Book stringsndjson() throws IOException {
        this.status = "strings.ndjson";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }

        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        Path dataDir = mdRoot.resolve("data");
        StringsNdjsonExporter.write(excelRoot, dataDir.resolve("strings.ndjson"));

        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir/md5/data} に各シートの {@code {sheet_id}.value.tsv} を書き出す（セル値は UTF-8 を Base64。
     * 空セルは空フィールド）。グリッドは {@code sheets.json} と同じ行・列サイズに合わせる。
     *
     * @see SheetValueTsvExporter
     */
    public Book sheetvalues() throws IOException {
        this.status = "sheet.value.tsv";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }

        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        Path dataDir = mdRoot.resolve("data");
        SheetValueTsvExporter.writeValueTsvFiles(excelRoot, dataDir);

        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir/md5/data} に各シートの {@code {sheet_id}.type.tsv} を書き出す（セル型はプレーンテキスト:
     * {@code string}, {@code number}, {@code boolean}, {@code error}, {@code date}。空セルは空フィールド）。
     */
    public Book sheettypes() throws IOException {
        this.status = "sheet.type.tsv";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }

        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        Path dataDir = mdRoot.resolve("data");
        SheetValueTsvExporter.writeTypeTsvFiles(excelRoot, dataDir);

        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir/md5/data} に各シートの {@code {sheet_id}.string.tsv} を書き出す（{@code t="s"} のセルのみ {@code strings.ndjson}
     * の {@code id} をプレーンテキスト）。
     *
     * @see SheetValueTsvExporter#writeStringTsvFiles
     */
    public Book sheetstrings() throws IOException {
        this.status = "sheet.string.tsv";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }

        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        Path dataDir = mdRoot.resolve("data");
        SheetValueTsvExporter.writeStringTsvFiles(excelRoot, dataDir);

        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir/md5/data} に各シートの {@code {sheet_id}.formula.tsv} を書き出す（{@code <f>} のプレーンテキスト。
     * 数式が無いセルは空フィールド）。
     */
    public Book sheetformulas() throws IOException {
        this.status = "sheet.formula.tsv";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }

        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        Path dataDir = mdRoot.resolve("data");
        SheetValueTsvExporter.writeFormulaTsvFiles(excelRoot, dataDir);

        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir/md5/data} に各シートの {@code {sheet_id}.format.tsv} を書き出す（{@code styles.xml} の {@code cellXf}
     * に対応する数値書式コード文字列。プレーンテキスト）。
     *
     * @see SheetValueTsvExporter#writeFormatTsvFiles
     */
    public Book sheetformats() throws IOException {
        this.status = "sheet.format.tsv";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }

        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        Path dataDir = mdRoot.resolve("data");
        SheetValueTsvExporter.writeFormatTsvFiles(excelRoot, dataDir);

        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir/md5/data} に各シートの {@code {sheet_id}.merge.tsv} を書き出す（結合セル。プレーンテキスト。
     * {@code docs/data-tsv-export.md} の merge 規約）。
     *
     * @see SheetValueTsvExporter#writeMergeTsvFiles
     */
    public Book sheetmerges() throws IOException {
        this.status = "sheet.merge.tsv";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }

        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        Path dataDir = mdRoot.resolve("data");
        SheetValueTsvExporter.writeMergeTsvFiles(excelRoot, dataDir);

        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir/md5/data} に各シートの {@code {sheet_id}.style.tsv} を書き出す（{@code cellXf} インデックスの参照 ID のみ。
     * プレーンテキスト）。
     *
     * @see SheetValueTsvExporter#writeStyleTsvFiles
     */
    public Book sheetstyles() throws IOException {
        this.status = "sheet.style.tsv";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }

        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        Path dataDir = mdRoot.resolve("data");
        SheetValueTsvExporter.writeStyleTsvFiles(excelRoot, dataDir);

        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir/md5/data/styles.ndjson} に {@code xl/styles.xml} の {@code cellXfs} を、{@code .style.tsv} の ID ごとの
     * レコード（NDJSON、1 行 1 オブジェクト）として書き出す。
     *
     * @see StylesJsonExporter#write
     */
    public Book stylesjson() throws IOException {
        this.status = "styles.ndjson";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }

        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        Path dataDir = mdRoot.resolve("data");
        StylesJsonExporter.write(excelRoot, dataDir.resolve("styles.ndjson"));

        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir/md5/data} に各シートの {@code {sheet_id}.formatted_value.tsv} を書き出す（簡易書式を適用した表示用文字列を
     * UTF-8 を Base64。未対応・失敗時は {@code value} と同じ文字列）。
     */
    public Book sheetformattedvalues() throws IOException {
        this.status = "sheet.formatted_value.tsv";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }

        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        Path dataDir = mdRoot.resolve("data");
        SheetValueTsvExporter.writeFormattedValueTsvFiles(excelRoot, dataDir);

        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir/md5/data} の TSV と {@code sheets.json} から、指定 {@code sheet_id} のワークシート XML を
     * 展開先 {@code excel} ルート上に再書き込みする（{@code load} 後に呼ぶ）。
     *
     * @see WorksheetTsvRebuilder#rebuildWorksheet
     */
    public Book rebuildworksheetfromtsv(String sheetId) throws IOException {
        this.status = "rebuild.worksheet.tsv";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }
        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        Path dataDir = mdRoot.resolve("data");
        WorksheetTsvRebuilder.rebuildWorksheet(excelRoot, dataDir, sheetId);
        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir/md5/data} の指定シートのグリッド TSV（{@code value} / {@code type} / {@code string} / {@code formula} /
     * {@code format} / {@code merge} / {@code style} / {@code formatted_value}）から、ワークシート上の<strong>複数の行を同時削除</strong>する。
     *
     * <p>削除行はすべて <strong>1 始まり</strong>。一時ファイルに書き換え結果を並べてからすべてのファイルで行・列サイズが揃っていることを確認したあとだけ、原本を削除してリネームする。
     * 処理が失敗した場合は作成した一時ファイルだけ削除し、元のデータは残る。
     *
     * @param sheetId {@code sheets.json} の {@code sheet_id}
     * @param rowsOneBased 削除対象となる行番号（1 始まり。重複可）
     */
    public Book datasheetdelrows(String sheetId, int... rowsOneBased) throws IOException {
        this.status = "data.sheet.tsv.del.rows";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (md5 == null || md5.isBlank()) {
            throw new IOException("MD5 が未定義です（先に load を実行してください）");
        }

        Path excelRoot = tmpdir.toPath().toAbsolutePath().normalize();
        Path mdRoot = excelRoot.getParent();
        if (mdRoot == null) {
            throw new IOException("作業ディレクトリの親を解決できません");
        }
        SheetTsvRowDeletion.deleteRows(mdRoot.resolve("data"), sheetId, rowsOneBased);
        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@link #outputdir} に、元ファイルと同名の xlsx を書き出す（{@code load} 後に呼ぶ）。
     */
    public Book zip() throws IOException {
        if (excelfile == null || outputdir == null) {
            throw new IOException("load が完了しておらず、出力先が未定義です");
        }
        return zip(new File(outputdir, excelfile.getName()));
    }

    /**
     * {@code tmpdir}（展開済み OOXML）を ZIP にまとめ、{@code destinationXlsx} に書き込む。
     */
    public Book zip(File destinationXlsx) throws IOException {
        this.status = "zip";
        if (tmpdir == null || !tmpdir.isDirectory()) {
            throw new IOException("展開ディレクトリがありません（先に load を実行してください）");
        }
        if (destinationXlsx == null) {
            throw new IOException("出力先ファイルが null です");
        }

        Path destPath = destinationXlsx.toPath().toAbsolutePath().normalize();
        Path parent = destPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        packZipArchive(tmpdir.toPath(), destPath);

        this.endTime = System.currentTimeMillis();
        return this;
    }

    /**
     * {@code tmp.dir}/{@link #md5} 以下の作業ディレクトリ（展開した {@code excel} を含む）を削除する。
     */
    public Book clear() throws IOException {
        this.status = "clear";
        if (md5 == null || md5.isBlank()) {
            this.endTime = System.currentTimeMillis();
            return this;
        }

        Properties props;
        try {
            props = ExternalConfig.load(Book.class);
        } catch (URISyntaxException e) {
            throw new IOException("設定の解決に失敗しました", e);
        }

        Path baseTmp = Path.of(props.getProperty("tmp.dir", "tmp")).toAbsolutePath().normalize();
        Path mdDir = baseTmp.resolve(md5);
        deleteRecursively(mdDir);

        this.tmpdir = null;
        this.endTime = System.currentTimeMillis();
        return this;
    }

    protected Book unzip() throws IOException {
        this.status = "unzip";

        Properties props;
        try {
            props = ExternalConfig.load(Book.class);
        } catch (URISyntaxException e) {
            throw new IOException("設定の解決に失敗しました", e);
        }

        Path baseTmp = Path.of(props.getProperty("tmp.dir", "tmp")).toAbsolutePath().normalize();
        Files.createDirectories(baseTmp);

        Path outputPath = Path.of(props.getProperty("output.dir", "output")).toAbsolutePath().normalize();
        Path logPath = Path.of(props.getProperty("log.dir", "log")).toAbsolutePath().normalize();
        this.outputdir = outputPath.toFile();
        this.logdir = logPath.toFile();
        Files.createDirectories(outputPath);
        Files.createDirectories(logPath);

        if (excelfile == null || !excelfile.isFile()) {
            throw new IOException("Excel ファイルが存在しません: " + excelfile);
        }

        String md5Hex = md5HexStreaming(excelfile.toPath());
        this.md5 = md5Hex;

        Path mdDir = baseTmp.resolve(md5Hex);
        deleteRecursively(mdDir);

        Path excelDir = mdDir.resolve("excel");
        Files.createDirectories(excelDir);

        unpackZipArchive(excelfile.toPath(), excelDir);

        this.tmpdir = excelDir.toFile();
        this.endTime = System.currentTimeMillis();
        return this;
    }

    private static String md5HexStreaming(Path file) throws IOException {
        MessageDigest digest;
        try {
            digest = MessageDigest.getInstance("MD5");
        } catch (NoSuchAlgorithmException e) {
            throw new IOException("MD5 が利用できません", e);
        }

        try (InputStream in = Files.newInputStream(file)) {
            byte[] buf = new byte[STREAM_BUFFER_SIZE];
            int n;
            while ((n = in.read(buf)) >= 0) {
                digest.update(buf, 0, n);
            }
        }

        return HexFormat.of().formatHex(digest.digest());
    }

    /** .xlsx（OOXML）は ZIP のため、指定ディレクトリへ展開する（Zip Slip を拒否）。 */
    private static void unpackZipArchive(Path xlsxFile, Path destDir) throws IOException {
        Path destNorm = destDir.toAbsolutePath().normalize();
        Files.createDirectories(destNorm);

        try (ZipFile zip = new ZipFile(xlsxFile.toFile())) {
            var entries = zip.stream().iterator();
            while (entries.hasNext()) {
                ZipEntry entry = entries.next();
                Path target = destNorm.resolve(entry.getName()).normalize();
                if (!target.startsWith(destNorm)) {
                    throw new IOException("ZIP エントリが展開先の外を指しています: " + entry.getName());
                }
                if (entry.isDirectory()) {
                    Files.createDirectories(target);
                } else {
                    Files.createDirectories(target.getParent());
                    try (InputStream in = zip.getInputStream(entry)) {
                        Files.copy(in, target, StandardCopyOption.REPLACE_EXISTING);
                    }
                }
            }
        }
    }

    /** ディレクトリ {@code sourceRoot} の内容を ZIP（エントリ名は {@code /} 区切り）として書き出す。 */
    private static void packZipArchive(Path sourceRoot, Path zipOutPath) throws IOException {
        Path rootNorm = sourceRoot.toAbsolutePath().normalize();

        try (OutputStream raw = Files.newOutputStream(zipOutPath);
             BufferedOutputStream buffered = new BufferedOutputStream(raw);
             ZipOutputStream zos = new ZipOutputStream(buffered)) {

            Files.walkFileTree(rootNorm, new SimpleFileVisitor<Path>() {
                @Override
                public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                    Path relative = rootNorm.relativize(file.normalize());
                    String entryName = relative.toString().replace('\\', '/');
                    zos.putNextEntry(new ZipEntry(entryName));
                    Files.copy(file, zos);
                    zos.closeEntry();
                    return FileVisitResult.CONTINUE;
                }
            });
        }
    }

    private static void deleteRecursively(Path root) throws IOException {
        if (!Files.exists(root)) {
            return;
        }
        Files.walkFileTree(root, new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                Files.delete(file);
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult postVisitDirectory(Path dir, IOException exc) throws IOException {
                if (exc != null) {
                    throw exc;
                }
                Files.delete(dir);
                return FileVisitResult.CONTINUE;
            }
        });
    }
}
