package jp.engawa.excel002;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.zip.ZipFile;

import org.junit.jupiter.api.Test;

class BookTest {

    /** {@code load} → {@code zip} → {@code clear} で出力と tmp 作業領域の削除まで確認する。 */
    @Test
    void unzipBook() throws Exception {
        URL url = BookTest.class.getResource("/input/sampledata.xlsx");
        assertNotNull(url, "classpath: input/sampledata.xlsx（テストリソース）が必要です");

        Path xlsx = Path.of(url.toURI());
        Book book = Book.load(xlsx.toFile());

        assertNotNull(book);
        assertNotNull(book.md5);
        assertTrue(Files.isRegularFile(book.tmpdir.toPath().resolve("[Content_Types].xml")),
                "展開結果に OOXML の [Content_Types].xml があること");

        Path mdWorkDir = book.tmpdir.toPath().toAbsolutePath().normalize().getParent();

        Path expectedOut = book.outputdir.toPath().resolve(xlsx.getFileName().toString());
        book.zip();

        assertTrue(Files.isRegularFile(expectedOut), "output に元と同じファイル名で xlsx が出力されていること");
        try (ZipFile z = new ZipFile(expectedOut.toFile())) {
            assertNotNull(z.getEntry("[Content_Types].xml"), "再 ZIP に [Content_Types].xml が含まれること");
        }

        book.clear();
        assertFalse(Files.exists(mdWorkDir), "tmp.dir/md5 の作業ディレクトリが削除されていること");
        assertNull(book.tmpdir);
    }
}
