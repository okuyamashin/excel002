package jp.engawa.excel002;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.net.URL;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import java.util.Base64;
import java.util.List;
import java.util.zip.ZipFile;

import org.junit.jupiter.api.Test;

class BookTest {

    private static void assertSameUtf8File(Path expected, Path actual, String relative) throws IOException {
        String e = Files.readString(expected.resolve(relative), StandardCharsets.UTF_8);
        String a = Files.readString(actual.resolve(relative), StandardCharsets.UTF_8);
        assertEquals(e, a, relative + " が再 zip → 再展開 → 再出力後も一致すること");
    }

    /** {@code load} → {@code zip} → {@code clear} で出力と tmp 作業領域の削除まで確認する。 */
    @Test
    void unzipBook() throws Exception {
        Book.clearAllTmpDir();
        
        URL url = BookTest.class.getResource("/input/sampledata.xlsx");
        assertNotNull(url, "classpath: input/sampledata.xlsx（テストリソース）が必要です");

        Path xlsx = Path.of(url.toURI());
        Book book = Book.load(xlsx.toFile());

        assertNotNull(book);
        assertNotNull(book.md5);
        assertTrue(Files.isRegularFile(book.tmpdir.toPath().resolve("[Content_Types].xml")),
                "展開結果に OOXML の [Content_Types].xml があること");

        Path mdWorkDir = book.tmpdir.toPath().toAbsolutePath().normalize().getParent();

        Path dataDir = mdWorkDir.resolve("data");
        assertTrue(Files.isRegularFile(dataDir.resolve("sheets.json")), "data/sheets.json が出力されていること");
        Path stringsNdjson = dataDir.resolve("strings.ndjson");
        assertTrue(Files.isRegularFile(stringsNdjson), "data/strings.ndjson が出力されていること");
        List<String> stringRecords =
                Files.readAllLines(stringsNdjson, StandardCharsets.UTF_8).stream()
                        .filter(line -> !line.isBlank())
                        .toList();
        assertEquals(10, stringRecords.size(), "サンプル共有文字列は 10 件（uniqueCount）");
        assertTrue(stringRecords.get(0).contains("\"schema_version\":1"), "strings NDJSON に schema_version");
        assertTrue(stringRecords.get(0).contains("\"id\":\"0\""), "先頭 id が 0");
        assertTrue(stringRecords.get(0).contains("sampledata"), "id 0 の value に sampledata");

        Path valueTsv = dataDir.resolve("1.value.tsv");
        assertTrue(Files.isRegularFile(valueTsv), "data/1.value.tsv が出力されていること");
        String firstRow = Files.readString(valueTsv, StandardCharsets.UTF_8).split("\n", 2)[0];
        String[] firstCols = firstRow.split("\t", -1);
        assertEquals(
                "sampledata",
                new String(Base64.getDecoder().decode(firstCols[0]), StandardCharsets.UTF_8),
                "A1 の値が Base64 でデコードできること");

        Path typeTsv = dataDir.resolve("1.type.tsv");
        assertTrue(Files.isRegularFile(typeTsv), "data/1.type.tsv が出力されていること");
        String typeFirstRow = Files.readString(typeTsv, StandardCharsets.UTF_8).split("\n", 2)[0];
        assertEquals("string", typeFirstRow.split("\t", -1)[0], "A1 の型が string であること");

        Path stringTsv = dataDir.resolve("1.string.tsv");
        assertTrue(Files.isRegularFile(stringTsv), "data/1.string.tsv が出力されていること");
        String[] stringRow1 = Files.readString(stringTsv, StandardCharsets.UTF_8).split("\n")[0].split("\t", -1);
        assertEquals("0", stringRow1[0], "A1 は共有文字列 id 0");
        assertEquals("", stringRow1[1], "B1 は共有文字列参照なしで空フィールド");
        Path stringTsv2 = dataDir.resolve("2.string.tsv");
        assertTrue(Files.isRegularFile(stringTsv2), "data/2.string.tsv が出力されていること");
        assertEquals(
                "1",
                Files.readString(stringTsv2, StandardCharsets.UTF_8).split("\n")[0].split("\t", -1)[0],
                "Sheet2 A1 は共有文字列 id 1（テストデータ依存）");

        Path formulaTsv = dataDir.resolve("1.formula.tsv");
        assertTrue(Files.isRegularFile(formulaTsv), "data/1.formula.tsv が出力されていること");
        String[] row2 = Files.readString(formulaTsv, StandardCharsets.UTF_8).split("\n")[1].split("\t", -1);
        assertEquals("D2+E2", row2[5], "F2 の数式が D2+E2 であること");

        Path formatTsv = dataDir.resolve("1.format.tsv");
        assertTrue(Files.isRegularFile(formatTsv), "data/1.format.tsv が出力されていること");
        String[] formatLines = Files.readString(formatTsv, StandardCharsets.UTF_8).split("\n");
        assertEquals(
                "General",
                formatLines[0].split("\t", -1)[0],
                "A1 の書式コードが cellXf に応じて General であること");
        assertEquals(
                "m/d/yyyy",
                formatLines[4].split("\t", -1)[0],
                "5 行目の日付列が組み込み書式 m/d/yyyy であること");

        Path mergeTsv = dataDir.resolve("1.merge.tsv");
        assertTrue(Files.isRegularFile(mergeTsv), "data/1.merge.tsv が出力されていること");
        String[] mergeRow1 = Files.readString(mergeTsv, StandardCharsets.UTF_8).split("\n")[0].split("\t", -1);
        assertEquals("0,1", mergeRow1[0], "結合の代表セル A1 が 0,1 であること");
        assertEquals("(1,1)", mergeRow1[1], "結合に飲み込まれた B1 が結合元 A1 を (列,行) で参照すること");
        assertEquals("0,0", mergeRow1[2], "非結合セル C1 が 0,0 であること");

        Path mergeSheet2 = dataDir.resolve("2.merge.tsv");
        assertTrue(Files.isRegularFile(mergeSheet2), "data/2.merge.tsv が出力されていること");
        assertTrue(
                Files.readString(mergeSheet2, StandardCharsets.UTF_8).contains("0,0"),
                "結合なしシートは 0,0 のみであること");

        Path styleTsv = dataDir.resolve("1.style.tsv");
        assertTrue(Files.isRegularFile(styleTsv), "data/1.style.tsv が出力されていること");
        String[] styleLines = Files.readString(styleTsv, StandardCharsets.UTF_8).split("\n");
        String[] styleRow1 = styleLines[0].split("\t", -1);
        assertEquals("4", styleRow1[0], "A1 の style は cellXf インデックス 4");
        assertEquals("4", styleRow1[1], "B1 の style は cellXf インデックス 4");
        assertEquals(
                "0",
                styleLines[1].split("\t", -1)[0],
                "s 省略セルは既定の cellXf 0");
        assertEquals(
                "1",
                styleLines[4].split("\t", -1)[0],
                "日付行の cellXf インデックス 1");

        Path stylesNdjson = dataDir.resolve("styles.ndjson");
        assertTrue(Files.isRegularFile(stylesNdjson), "data/styles.ndjson が出力されていること");
        List<String> xfsNdLines =
                Files.readAllLines(stylesNdjson, StandardCharsets.UTF_8).stream()
                        .filter(line -> !line.isBlank())
                        .toList();
        assertEquals(5, xfsNdLines.size(), "サンプルブックの cellXfs は 5 件");
        assertTrue(xfsNdLines.get(0).contains("\"schema_version\":1"), "schema_version があること");
        assertTrue(xfsNdLines.get(0).contains("\"id\":\"0\""), "先頭行が id 0");
        assertTrue(xfsNdLines.get(4).contains("\"id\":\"4\""), "id 4 のレコードがあること");
        assertTrue(xfsNdLines.get(0).contains("\"restore_xml\""), "restore_xml キーがあること");
        assertTrue(
                String.join("\n", xfsNdLines).contains("spreadsheetml/2006/main"),
                "restore_xml にスタイルシートの名前空間 URI が含まれること");

        Path formattedTsv = dataDir.resolve("1.formatted_value.tsv");
        assertTrue(Files.isRegularFile(formattedTsv), "data/1.formatted_value.tsv が出力されていること");
        String[] fvLines = Files.readString(formattedTsv, StandardCharsets.UTF_8).split("\n");
        assertEquals(
                "sampledata",
                new String(Base64.getDecoder().decode(fvLines[0].split("\t", -1)[0]), StandardCharsets.UTF_8),
                "formatted_value の文字列セルは value と同じであること");
        String dateShown =
                new String(Base64.getDecoder().decode(fvLines[4].split("\t", -1)[0]), StandardCharsets.UTF_8);
        assertTrue(
                dateShown.matches("\\d{4}-\\d{2}-\\d{2}.*"),
                "日付書式セルが yyyy-MM-dd 形式になること: " + dateShown);

        Path sheetXml = book.tmpdir.toPath().resolve("xl/worksheets/sheet1.xml");
        String xmlBefore = Files.readString(sheetXml, StandardCharsets.UTF_8);
        assertTrue(xmlBefore.contains("D2+E2"), "元の sheet1 に数式があること");

        WorksheetTsvRebuilder.rebuildWorksheet(book.tmpdir.toPath(), dataDir, "1");
        String xmlAfter = Files.readString(sheetXml, StandardCharsets.UTF_8);
        assertTrue(
                xmlAfter.contains("ref=\"A1:F7\""),
                "再構築後 dimension が A1:F7 であること: " + xmlAfter);
        assertTrue(xmlAfter.contains("D2+E2"), "再構築後も F2 の数式が残ること");
        assertTrue(xmlAfter.contains("D4+E4"), "再構築後も F4 の数式が残ること");
        assertTrue(
                xmlAfter.contains("ref=\"A1:B1\""),
                "再構築後 mergeCell が A1:B1 であること");
        book.rebuildworksheetfromtsv("2");
        Path sheet2 = book.tmpdir.toPath().resolve("xl/worksheets/sheet2.xml");
        String xml2 = Files.readString(sheet2, StandardCharsets.UTF_8);
        assertTrue(xml2.contains("sheetData"), "Sheet2 の再構築で sheetData があること");

        Path expectedOut = book.outputdir.toPath().resolve(xlsx.getFileName().toString());
        book.zip();

        assertTrue(Files.isRegularFile(expectedOut), "output に元と同じファイル名で xlsx が出力されていること");
        try (ZipFile z = new ZipFile(expectedOut.toFile())) {
            assertNotNull(z.getEntry("[Content_Types].xml"), "再 ZIP に [Content_Types].xml が含まれること");
        }

        //book.clear();
        //assertFalse(Files.exists(mdWorkDir), "tmp.dir/md5 の作業ディレクトリが削除されていること");
        //assertNull(book.tmpdir);
    }

    /**
     * 展開 → TSV から全シートの worksheet 再構築 → zip → 再 load したとき、セル由来の TSV が元と一致すること（実質の同一データ確認）。
     */
    @Test
    void loadRebuildWorksheetsZip_roundTripPreservesCellTsvExport() throws Exception {
        URL url = BookTest.class.getResource("/input/sampledata.xlsx");
        assertNotNull(url, "classpath: input/sampledata.xlsx が必要です");

        Path xlsx = Path.of(url.toURI());
        Book book = Book.load(xlsx.toFile());
        Path mdWorkDir = book.tmpdir.toPath().toAbsolutePath().normalize().getParent();
        Path dataDirOrig = mdWorkDir.resolve("data");

        book.rebuildworksheetfromtsv("1");
        book.rebuildworksheetfromtsv("2");

        Path roundtripXlsx = mdWorkDir.resolve("roundtrip-cell-check.xlsx");
        Files.deleteIfExists(roundtripXlsx);
        book.zip(roundtripXlsx.toFile());

        Book book2 = Book.load(roundtripXlsx.toFile());
        Path dataDirRound = book2.tmpdir.toPath().toAbsolutePath().normalize().getParent().resolve("data");

        for (String sid : List.of("1", "2")) {
            for (String aspect :
                    List.of("value.tsv", "type.tsv", "string.tsv", "formula.tsv", "merge.tsv", "style.tsv", "format.tsv")) {
                assertSameUtf8File(dataDirOrig, dataDirRound, sid + "." + aspect);
            }
        }
        assertSameUtf8File(dataDirOrig, dataDirRound, "strings.ndjson");
        assertSameUtf8File(dataDirOrig, dataDirRound, "styles.ndjson");
        assertSameUtf8File(dataDirOrig, dataDirRound, "1.formatted_value.tsv");
        assertSameUtf8File(dataDirOrig, dataDirRound, "2.formatted_value.tsv");

        Files.deleteIfExists(roundtripXlsx);
    }
}
