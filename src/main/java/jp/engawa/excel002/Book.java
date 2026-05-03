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
        return book;
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
