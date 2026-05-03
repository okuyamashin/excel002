package jp.engawa.excel002;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Base64;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

/**
 * {@code data/{sheet_id}.value.tsv} / {@code .type.tsv} / {@code .string.tsv} / {@code .formula.tsv} /
 * {@code .format.tsv} / {@code .merge.tsv} / {@code .style.tsv} / {@code .formatted_value.tsv} を書き出す（{@code docs/data-tsv-export.md}）。
 */
final class SheetValueTsvExporter {
    private static final String NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static final String NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static final String NS_REL_OD = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static final Pattern DIMENSION_REF =
            Pattern.compile("<dimension[^>]*ref=\"([^\"]*)\"", Pattern.CASE_INSENSITIVE);
    private static final Pattern CELL_R_ATTR =
            Pattern.compile("<c[^>]*\\br=\"([A-Za-z]+)(\\d+)\"");

    private SheetValueTsvExporter() {}

    private record CellRead(
            String valueText,
            String typeLabel,
            String formulaText,
            int styleIndex,
            Integer sharedStringIndex) {}

    /** {@code dataDir} に各シートの {@code {sheet_id}.value.tsv} を UTF-8・LF で出力する。 */
    static void writeValueTsvFiles(Path excelRoot, Path dataDir) throws IOException {
        Path xl = excelRoot.resolve("xl");
        if (!Files.isDirectory(xl)) {
            throw new IOException("xl ディレクトリがありません: " + excelRoot);
        }

        Path excelNorm = excelRoot.toAbsolutePath().normalize();
        List<String> sharedStrings = readSharedStrings(xl.resolve("sharedStrings.xml"));
        Map<String, String> worksheetTargets = readWorksheetTargets(xl.resolve("_rels/workbook.xml.rels"));
        List<ValueSheetRef> sheets =
                readWorksheetSheetRefs(xl.resolve("workbook.xml"), worksheetTargets, xl, excelNorm);

        Files.createDirectories(dataDir);

        for (ValueSheetRef sh : sheets) {
            Path worksheetFile = excelNorm.resolve(sh.worksheetRelativePath()).normalize();
            if (!worksheetFile.startsWith(excelNorm)) {
                throw new IOException("ワークシートパスが excel ルートの外です: " + worksheetFile);
            }
            IntBounds dim = worksheetUsedBounds(worksheetFile);
            CellRead[][] cells = loadCellReads(worksheetFile, sharedStrings, dim.rows(), dim.cols());
            Path out = dataDir.resolve(sh.sheetId() + ".value.tsv");
            writeGridBase64Tsv(out, toValueGrid(cells));
        }
    }

    /**
     * {@code dataDir} に各シートの {@code {sheet_id}.type.tsv} を UTF-8・LF で出力する（プレーンテキスト、タブ不可文字は含めない想定）。
     */
    static void writeTypeTsvFiles(Path excelRoot, Path dataDir) throws IOException {
        Path xl = excelRoot.resolve("xl");
        if (!Files.isDirectory(xl)) {
            throw new IOException("xl ディレクトリがありません: " + excelRoot);
        }

        Path excelNorm = excelRoot.toAbsolutePath().normalize();
        List<String> sharedStrings = readSharedStrings(xl.resolve("sharedStrings.xml"));
        Map<String, String> worksheetTargets = readWorksheetTargets(xl.resolve("_rels/workbook.xml.rels"));
        List<ValueSheetRef> sheets =
                readWorksheetSheetRefs(xl.resolve("workbook.xml"), worksheetTargets, xl, excelNorm);

        Files.createDirectories(dataDir);

        for (ValueSheetRef sh : sheets) {
            Path worksheetFile = excelNorm.resolve(sh.worksheetRelativePath()).normalize();
            if (!worksheetFile.startsWith(excelNorm)) {
                throw new IOException("ワークシートパスが excel ルートの外です: " + worksheetFile);
            }
            IntBounds dim = worksheetUsedBounds(worksheetFile);
            CellRead[][] cells = loadCellReads(worksheetFile, sharedStrings, dim.rows(), dim.cols());
            Path out = dataDir.resolve(sh.sheetId() + ".type.tsv");
            writeGridPlainTsv(out, toTypeGrid(cells));
        }
    }

    /**
     * {@code dataDir} に各シートの {@code {sheet_id}.string.tsv} を出力する（{@code t="s"} のセルのみ
     * {@code strings.ndjson} の {@code id} を十進文字列で。それ以外・{@code <c>} 無しは空フィールド）。
     */
    static void writeStringTsvFiles(Path excelRoot, Path dataDir) throws IOException {
        Path xl = excelRoot.resolve("xl");
        if (!Files.isDirectory(xl)) {
            throw new IOException("xl ディレクトリがありません: " + excelRoot);
        }

        Path excelNorm = excelRoot.toAbsolutePath().normalize();
        List<String> sharedStrings = readSharedStrings(xl.resolve("sharedStrings.xml"));
        Map<String, String> worksheetTargets = readWorksheetTargets(xl.resolve("_rels/workbook.xml.rels"));
        List<ValueSheetRef> sheets =
                readWorksheetSheetRefs(xl.resolve("workbook.xml"), worksheetTargets, xl, excelNorm);

        Files.createDirectories(dataDir);

        for (ValueSheetRef sh : sheets) {
            Path worksheetFile = excelNorm.resolve(sh.worksheetRelativePath()).normalize();
            if (!worksheetFile.startsWith(excelNorm)) {
                throw new IOException("ワークシートパスが excel ルートの外です: " + worksheetFile);
            }
            IntBounds dim = worksheetUsedBounds(worksheetFile);
            CellRead[][] cells = loadCellReads(worksheetFile, sharedStrings, dim.rows(), dim.cols());
            Path out = dataDir.resolve(sh.sheetId() + ".string.tsv");
            writeGridPlainTsv(out, toSharedStringIdGrid(cells));
        }
    }

    /**
     * {@code dataDir} に各シートの {@code {sheet_id}.formula.tsv} を UTF-8・LF で出力する（{@code <f>} のテキストをプレーンテキスト。式が無いセルは空フィールド）。
     */
    static void writeFormulaTsvFiles(Path excelRoot, Path dataDir) throws IOException {
        Path xl = excelRoot.resolve("xl");
        if (!Files.isDirectory(xl)) {
            throw new IOException("xl ディレクトリがありません: " + excelRoot);
        }

        Path excelNorm = excelRoot.toAbsolutePath().normalize();
        List<String> sharedStrings = readSharedStrings(xl.resolve("sharedStrings.xml"));
        Map<String, String> worksheetTargets = readWorksheetTargets(xl.resolve("_rels/workbook.xml.rels"));
        List<ValueSheetRef> sheets =
                readWorksheetSheetRefs(xl.resolve("workbook.xml"), worksheetTargets, xl, excelNorm);

        Files.createDirectories(dataDir);

        for (ValueSheetRef sh : sheets) {
            Path worksheetFile = excelNorm.resolve(sh.worksheetRelativePath()).normalize();
            if (!worksheetFile.startsWith(excelNorm)) {
                throw new IOException("ワークシートパスが excel ルートの外です: " + worksheetFile);
            }
            IntBounds dim = worksheetUsedBounds(worksheetFile);
            CellRead[][] cells = loadCellReads(worksheetFile, sharedStrings, dim.rows(), dim.cols());
            Path out = dataDir.resolve(sh.sheetId() + ".formula.tsv");
            writeGridPlainTsv(out, toFormulaGrid(cells));
        }
    }

    /**
     * {@code dataDir} に各シートの {@code {sheet_id}.format.tsv} を出力する（セルの {@code cellXf} に対応する数値書式コード。
     * プレーンテキスト・タブ改行は空白に置換。{@code <c>} が無いセルは空フィールド）。
     */
    static void writeFormatTsvFiles(Path excelRoot, Path dataDir) throws IOException {
        Path xl = excelRoot.resolve("xl");
        if (!Files.isDirectory(xl)) {
            throw new IOException("xl ディレクトリがありません: " + excelRoot);
        }

        Path excelNorm = excelRoot.toAbsolutePath().normalize();
        List<String> sharedStrings = readSharedStrings(xl.resolve("sharedStrings.xml"));
        Map<String, String> worksheetTargets = readWorksheetTargets(xl.resolve("_rels/workbook.xml.rels"));
        List<ValueSheetRef> sheets =
                readWorksheetSheetRefs(xl.resolve("workbook.xml"), worksheetTargets, xl, excelNorm);
        OoxmlStylesResolver styles = OoxmlStylesResolver.load(xl.resolve("styles.xml"));

        Files.createDirectories(dataDir);

        for (ValueSheetRef sh : sheets) {
            Path worksheetFile = excelNorm.resolve(sh.worksheetRelativePath()).normalize();
            if (!worksheetFile.startsWith(excelNorm)) {
                throw new IOException("ワークシートパスが excel ルートの外です: " + worksheetFile);
            }
            IntBounds dim = worksheetUsedBounds(worksheetFile);
            CellRead[][] cells = loadCellReads(worksheetFile, sharedStrings, dim.rows(), dim.cols());
            Path out = dataDir.resolve(sh.sheetId() + ".format.tsv");
            writeGridPlainTsv(out, toFormatGrid(cells, styles));
        }
    }

    /**
     * {@code dataDir} に各シートの {@code {sheet_id}.merge.tsv} を出力する（結合セル。プレーンテキスト。
     * 通常セル {@code 0,0}、結合の左上（代表）{@code 0,1}、結合に飲み込まれたセル {@code -1,-1}）。
     */
    static void writeMergeTsvFiles(Path excelRoot, Path dataDir) throws IOException {
        Path xl = excelRoot.resolve("xl");
        if (!Files.isDirectory(xl)) {
            throw new IOException("xl ディレクトリがありません: " + excelRoot);
        }

        Path excelNorm = excelRoot.toAbsolutePath().normalize();
        Map<String, String> worksheetTargets = readWorksheetTargets(xl.resolve("_rels/workbook.xml.rels"));
        List<ValueSheetRef> sheets =
                readWorksheetSheetRefs(xl.resolve("workbook.xml"), worksheetTargets, xl, excelNorm);

        Files.createDirectories(dataDir);

        for (ValueSheetRef sh : sheets) {
            Path worksheetFile = excelNorm.resolve(sh.worksheetRelativePath()).normalize();
            if (!worksheetFile.startsWith(excelNorm)) {
                throw new IOException("ワークシートパスが excel ルートの外です: " + worksheetFile);
            }
            IntBounds dim = worksheetUsedBounds(worksheetFile);
            List<MergeRect> merges = readMergeRects(worksheetFile);
            Path out = dataDir.resolve(sh.sheetId() + ".merge.tsv");
            writeGridPlainTsv(out, toMergeGrid(dim.rows(), dim.cols(), merges));
        }
    }

    /**
     * {@code dataDir} に各シートの {@code {sheet_id}.style.tsv} を出力する（{@code <c s="…">} の {@code cellXf}
     * インデックスを十進文字列の参照 ID のみ。{@code s} 省略時は {@code 0}。{@code <c>} が無いセルは空フィールド）。
     */
    static void writeStyleTsvFiles(Path excelRoot, Path dataDir) throws IOException {
        Path xl = excelRoot.resolve("xl");
        if (!Files.isDirectory(xl)) {
            throw new IOException("xl ディレクトリがありません: " + excelRoot);
        }

        Path excelNorm = excelRoot.toAbsolutePath().normalize();
        List<String> sharedStrings = readSharedStrings(xl.resolve("sharedStrings.xml"));
        Map<String, String> worksheetTargets = readWorksheetTargets(xl.resolve("_rels/workbook.xml.rels"));
        List<ValueSheetRef> sheets =
                readWorksheetSheetRefs(xl.resolve("workbook.xml"), worksheetTargets, xl, excelNorm);

        Files.createDirectories(dataDir);

        for (ValueSheetRef sh : sheets) {
            Path worksheetFile = excelNorm.resolve(sh.worksheetRelativePath()).normalize();
            if (!worksheetFile.startsWith(excelNorm)) {
                throw new IOException("ワークシートパスが excel ルートの外です: " + worksheetFile);
            }
            IntBounds dim = worksheetUsedBounds(worksheetFile);
            CellRead[][] cells = loadCellReads(worksheetFile, sharedStrings, dim.rows(), dim.cols());
            Path out = dataDir.resolve(sh.sheetId() + ".style.tsv");
            writeGridPlainTsv(out, toStyleGrid(cells));
        }
    }

    /**
     * {@code dataDir} に各シートの {@code {sheet_id}.formatted_value.tsv} を出力する（UTF-8・LF・セルは Base64。
     * 書式は簡易実装。失敗時は {@code value} と同じ文字列）。
     */
    static void writeFormattedValueTsvFiles(Path excelRoot, Path dataDir) throws IOException {
        Path xl = excelRoot.resolve("xl");
        if (!Files.isDirectory(xl)) {
            throw new IOException("xl ディレクトリがありません: " + excelRoot);
        }

        Path excelNorm = excelRoot.toAbsolutePath().normalize();
        List<String> sharedStrings = readSharedStrings(xl.resolve("sharedStrings.xml"));
        Map<String, String> worksheetTargets = readWorksheetTargets(xl.resolve("_rels/workbook.xml.rels"));
        List<ValueSheetRef> sheets =
                readWorksheetSheetRefs(xl.resolve("workbook.xml"), worksheetTargets, xl, excelNorm);
        OoxmlStylesResolver styles = OoxmlStylesResolver.load(xl.resolve("styles.xml"));

        Files.createDirectories(dataDir);

        for (ValueSheetRef sh : sheets) {
            Path worksheetFile = excelNorm.resolve(sh.worksheetRelativePath()).normalize();
            if (!worksheetFile.startsWith(excelNorm)) {
                throw new IOException("ワークシートパスが excel ルートの外です: " + worksheetFile);
            }
            IntBounds dim = worksheetUsedBounds(worksheetFile);
            CellRead[][] cells = loadCellReads(worksheetFile, sharedStrings, dim.rows(), dim.cols());
            Path out = dataDir.resolve(sh.sheetId() + ".formatted_value.tsv");
            writeGridBase64Tsv(out, toFormattedValueGrid(cells, styles));
        }
    }

    private record ValueSheetRef(String sheetId, String worksheetRelativePath) {}

    /** 結合範囲（1 始まり・両端含む。左上が代表セル）。 */
    private record MergeRect(int row1, int col1, int row2, int col2) {}

    private record IntBounds(int rows, int cols) {}

    private static List<ValueSheetRef> readWorksheetSheetRefs(
            Path workbookXml,
            Map<String, String> worksheetTargets,
            Path xlDir,
            Path excelRootNorm)
            throws IOException {
        if (!Files.isRegularFile(workbookXml)) {
            throw new IOException("workbook.xml がありません: " + workbookXml);
        }
        List<ValueSheetRef> list = new ArrayList<>();
        XMLInputFactory f = XMLInputFactory.newInstance();
        try (InputStream in = Files.newInputStream(workbookXml)) {
            XMLStreamReader r = f.createXMLStreamReader(in);
            while (r.hasNext()) {
                if (r.next() == XMLStreamConstants.START_ELEMENT
                        && "sheet".equals(r.getLocalName())
                        && NS_MAIN.equals(r.getNamespaceURI())) {
                    String sheetId = r.getAttributeValue(null, "sheetId");
                    String rid = r.getAttributeValue(NS_REL_OD, "id");
                    if (sheetId == null || rid == null) {
                        continue;
                    }
                    String target = worksheetTargets.get(rid);
                    if (target == null) {
                        continue;
                    }
                    Path sheetPath = xlDir.resolve(target).normalize().toAbsolutePath().normalize();
                    Path xlNorm = xlDir.toAbsolutePath().normalize();
                    if (!sheetPath.startsWith(xlNorm) || !sheetPath.startsWith(excelRootNorm)) {
                        throw new IOException("シートパスが不正です: " + target);
                    }
                    if (!Files.isRegularFile(sheetPath)) {
                        throw new IOException("ワークシート XML が見つかりません: " + sheetPath);
                    }
                    Path rel = excelRootNorm.relativize(sheetPath);
                    if (rel.isAbsolute()) {
                        throw new IOException("ワークシートの相対パスを解決できません: " + sheetPath);
                    }
                    list.add(new ValueSheetRef(sheetId, rel.toString().replace('\\', '/')));
                }
            }
            r.close();
        } catch (XMLStreamException e) {
            throw new IOException("workbook.xml の解析に失敗しました", e);
        }
        return list;
    }

    private static Map<String, String> readWorksheetTargets(Path relsFile) throws IOException {
        if (!Files.isRegularFile(relsFile)) {
            throw new IOException("workbook.xml.rels がありません: " + relsFile);
        }
        Map<String, String> map = new HashMap<>();
        XMLInputFactory f = XMLInputFactory.newInstance();
        try (InputStream in = Files.newInputStream(relsFile)) {
            XMLStreamReader r = f.createXMLStreamReader(in);
            while (r.hasNext()) {
                if (r.next() == XMLStreamConstants.START_ELEMENT
                        && "Relationship".equals(r.getLocalName())
                        && NS_REL_PKG.equals(r.getNamespaceURI())) {
                    String id = r.getAttributeValue(null, "Id");
                    String target = r.getAttributeValue(null, "Target");
                    String type = r.getAttributeValue(null, "Type");
                    if (id != null
                            && target != null
                            && type != null
                            && type.contains("worksheet")) {
                        String norm = target.replace('\\', '/');
                        if (norm.startsWith("/")) {
                            norm = norm.substring(1);
                        }
                        map.put(id, norm);
                    }
                }
            }
            r.close();
        } catch (XMLStreamException e) {
            throw new IOException("workbook.xml.rels の解析に失敗しました", e);
        }
        return map;
    }

    static List<String> readSharedStrings(Path sharedXml) throws IOException {
        if (!Files.isRegularFile(sharedXml)) {
            return List.of();
        }
        List<String> strings = new ArrayList<>();
        XMLInputFactory f = XMLInputFactory.newInstance();
        try (InputStream in = Files.newInputStream(sharedXml)) {
            XMLStreamReader r = f.createXMLStreamReader(in);
            while (r.hasNext()) {
                if (r.next() == XMLStreamConstants.START_ELEMENT
                        && "si".equals(r.getLocalName())) {
                    strings.add(readSiPlainText(r));
                }
            }
            r.close();
        } catch (XMLStreamException e) {
            throw new IOException("sharedStrings.xml の解析に失敗しました", e);
        }
        return strings;
    }

    private static String readSiPlainText(XMLStreamReader r) throws XMLStreamException {
        StringBuilder sb = new StringBuilder();
        int depth = 1;
        while (r.hasNext() && depth > 0) {
            int ev = r.next();
            if (ev == XMLStreamConstants.START_ELEMENT && "t".equals(r.getLocalName())) {
                sb.append(r.getElementText());
            } else if (ev == XMLStreamConstants.START_ELEMENT) {
                depth++;
            } else if (ev == XMLStreamConstants.END_ELEMENT) {
                depth--;
            }
        }
        return sb.toString();
    }

    /** 使用範囲の行数・列数（いずれも 1 始まりの最大行・列と一致するサイズ）。無い場合は 0。 */
    private static IntBounds worksheetUsedBounds(Path worksheetXmlFile) throws IOException {
        String xml = Files.readString(worksheetXmlFile, StandardCharsets.UTF_8);
        Matcher dm = DIMENSION_REF.matcher(xml);
        if (dm.find()) {
            String ref = dm.group(1).trim();
            if (!ref.isEmpty()) {
                IntBounds b = boundsSizeFromRef(ref);
                if (b.rows() > 0 && b.cols() > 0) {
                    return b;
                }
            }
        }
        int maxRow = 0;
        int maxCol = 0;
        Matcher cm = CELL_R_ATTR.matcher(xml);
        while (cm.find()) {
            maxRow = Math.max(maxRow, Integer.parseInt(cm.group(2)));
            maxCol = Math.max(maxCol, columnLettersOneBased(cm.group(1)));
        }
        return new IntBounds(maxRow, maxCol);
    }

    private static IntBounds boundsSizeFromRef(String rangeRef) {
        String[] parts = rangeRef.split(":");
        if (parts.length == 1) {
            RowCol rc = parseCellOneBased(parts[0]);
            return new IntBounds(rc.row(), rc.col());
        }
        RowCol a = parseCellOneBased(parts[0]);
        RowCol b = parseCellOneBased(parts[1]);
        int maxRow = Math.max(a.row(), b.row());
        int maxCol = Math.max(a.col(), b.col());
        return new IntBounds(maxRow, maxCol);
    }

    private record RowCol(int row, int col) {}

    private static RowCol parseCellOneBased(String cell) {
        int i = 0;
        while (i < cell.length() && Character.isLetter(cell.charAt(i))) {
            i++;
        }
        if (i == 0 || i >= cell.length()) {
            return new RowCol(0, 0);
        }
        String letters = cell.substring(0, i).toUpperCase();
        int row = Integer.parseInt(cell.substring(i));
        int col = columnLettersOneBased(letters);
        return new RowCol(row, col);
    }

    private static int columnLettersOneBased(String letters) {
        int col = 0;
        for (int k = 0; k < letters.length(); k++) {
            col = col * 26 + (letters.charAt(k) - 'A' + 1);
        }
        return col;
    }

    private static List<MergeRect> readMergeRects(Path worksheetXml) throws IOException {
        List<MergeRect> list = new ArrayList<>();
        XMLInputFactory f = XMLInputFactory.newInstance();
        try (InputStream in = Files.newInputStream(worksheetXml)) {
            XMLStreamReader r = f.createXMLStreamReader(in);
            while (r.hasNext()) {
                if (r.next() != XMLStreamConstants.START_ELEMENT) {
                    continue;
                }
                if (!NS_MAIN.equals(r.getNamespaceURI()) || !"mergeCell".equals(r.getLocalName())) {
                    continue;
                }
                String ref = r.getAttributeValue(null, "ref");
                if (ref == null || ref.isBlank()) {
                    continue;
                }
                MergeRect rect = parseMergeRef(ref.trim());
                if (rect != null) {
                    list.add(rect);
                }
            }
            r.close();
        } catch (XMLStreamException e) {
            throw new IOException("ワークシートの mergeCells の解析に失敗しました: " + worksheetXml, e);
        }
        return list;
    }

    /** {@code A1} または {@code A1:B2}。不正なら {@code null}。 */
    private static MergeRect parseMergeRef(String ref) {
        String[] parts = ref.split(":");
        RowCol a = parseCellOneBased(parts[0].trim());
        RowCol b = parts.length == 1 ? a : parseCellOneBased(parts[1].trim());
        if (a.row() <= 0 || a.col() <= 0 || b.row() <= 0 || b.col() <= 0) {
            return null;
        }
        int r1 = Math.min(a.row(), b.row());
        int r2 = Math.max(a.row(), b.row());
        int c1 = Math.min(a.col(), b.col());
        int c2 = Math.max(a.col(), b.col());
        return new MergeRect(r1, c1, r2, c2);
    }

    private static String[][] toMergeGrid(int rowCount, int colCount, List<MergeRect> merges) {
        if (rowCount <= 0 || colCount <= 0) {
            return new String[0][0];
        }
        String[][] g = new String[rowCount][colCount];
        for (int r = 0; r < rowCount; r++) {
            for (int c = 0; c < colCount; c++) {
                g[r][c] = "0,0";
            }
        }
        for (MergeRect m : merges) {
            int r0Top = m.row1() - 1;
            int c0Left = m.col1() - 1;
            int r0Bot = m.row2() - 1;
            int c0Right = m.col2() - 1;
            int rStart = Math.max(0, r0Top);
            int rEnd = Math.min(rowCount - 1, r0Bot);
            int cStart = Math.max(0, c0Left);
            int cEnd = Math.min(colCount - 1, c0Right);
            if (rStart > rEnd || cStart > cEnd) {
                continue;
            }
            for (int r = rStart; r <= rEnd; r++) {
                for (int c = cStart; c <= cEnd; c++) {
                    if (r == r0Top && c == c0Left) {
                        g[r][c] = "0,1";
                    } else {
                        g[r][c] = "-1,-1";
                    }
                }
            }
        }
        return g;
    }

    private static CellRead[][] loadCellReads(
            Path worksheetXml, List<String> sharedStrings, int rowCount, int colCount) throws IOException {
        if (rowCount <= 0 || colCount <= 0) {
            return new CellRead[0][0];
        }
        CellRead[][] grid = new CellRead[rowCount][colCount];
        XMLInputFactory f = XMLInputFactory.newInstance();
        try (InputStream in = Files.newInputStream(worksheetXml)) {
            XMLStreamReader r = f.createXMLStreamReader(in);
            boolean inSheetData = false;
            while (r.hasNext()) {
                int ev = r.next();
                if (ev == XMLStreamConstants.START_ELEMENT && "sheetData".equals(r.getLocalName())) {
                    inSheetData = true;
                } else if (ev == XMLStreamConstants.END_ELEMENT && "sheetData".equals(r.getLocalName())) {
                    inSheetData = false;
                } else if (inSheetData
                        && ev == XMLStreamConstants.START_ELEMENT
                        && "c".equals(r.getLocalName())) {
                    String ref = r.getAttributeValue(null, "r");
                    String t = r.getAttributeValue(null, "t");
                    String sAttr = r.getAttributeValue(null, "s");
                    int styleIdx = parseCellStyleIndex(sAttr);
                    if (ref == null) {
                        consumeCellSubtree(r);
                        continue;
                    }
                    RowCol rc = parseCellOneBased(ref);
                    int row0 = rc.row() - 1;
                    int col0 = rc.col() - 1;
                    if (row0 < 0 || col0 < 0 || row0 >= rowCount || col0 >= colCount) {
                        consumeCellSubtree(r);
                        continue;
                    }
                    grid[row0][col0] = readCellPayload(r, t, sharedStrings, styleIdx);
                }
            }
            r.close();
        } catch (XMLStreamException e) {
            throw new IOException("ワークシートの解析に失敗しました: " + worksheetXml, e);
        }
        return grid;
    }

    private static int parseCellStyleIndex(String sAttr) {
        if (sAttr == null || sAttr.isBlank()) {
            return -1;
        }
        try {
            return Integer.parseInt(sAttr.trim());
        } catch (NumberFormatException e) {
            return -1;
        }
    }

    private static String[][] toValueGrid(CellRead[][] cells) {
        String[][] g = new String[cells.length][];
        for (int r = 0; r < cells.length; r++) {
            g[r] = new String[cells[r].length];
            for (int c = 0; c < cells[r].length; c++) {
                CellRead cr = cells[r][c];
                if (cr != null && cr.valueText() != null && !cr.valueText().isEmpty()) {
                    g[r][c] = cr.valueText();
                }
            }
        }
        return g;
    }

    private static String[][] toTypeGrid(CellRead[][] cells) {
        String[][] g = new String[cells.length][];
        for (int r = 0; r < cells.length; r++) {
            g[r] = new String[cells[r].length];
            for (int c = 0; c < cells[r].length; c++) {
                CellRead cr = cells[r][c];
                if (cr != null && cr.typeLabel() != null && !cr.typeLabel().isEmpty()) {
                    g[r][c] = cr.typeLabel();
                }
            }
        }
        return g;
    }

    private static String[][] toFormulaGrid(CellRead[][] cells) {
        String[][] g = new String[cells.length][];
        for (int r = 0; r < cells.length; r++) {
            g[r] = new String[cells[r].length];
            for (int c = 0; c < cells[r].length; c++) {
                CellRead cr = cells[r][c];
                if (cr != null && cr.formulaText() != null && !cr.formulaText().isEmpty()) {
                    g[r][c] = cr.formulaText();
                }
            }
        }
        return g;
    }

    private static String[][] toSharedStringIdGrid(CellRead[][] cells) {
        String[][] g = new String[cells.length][];
        for (int r = 0; r < cells.length; r++) {
            g[r] = new String[cells[r].length];
            for (int c = 0; c < cells[r].length; c++) {
                CellRead cr = cells[r][c];
                if (cr == null) {
                    continue;
                }
                Integer idx = cr.sharedStringIndex();
                if (idx != null) {
                    g[r][c] = Integer.toString(idx);
                }
            }
        }
        return g;
    }

    private static String[][] toFormattedValueGrid(CellRead[][] cells, OoxmlStylesResolver styles) {
        String[][] g = new String[cells.length][];
        for (int r = 0; r < cells.length; r++) {
            g[r] = new String[cells[r].length];
            for (int c = 0; c < cells[r].length; c++) {
                CellRead cr = cells[r][c];
                if (cr == null) {
                    continue;
                }
                String out =
                        SimpleFormattedValue.format(
                                cr.valueText(), cr.typeLabel(), cr.styleIndex(), styles);
                if (out != null && !out.isEmpty()) {
                    g[r][c] = out;
                }
            }
        }
        return g;
    }

    private static String[][] toFormatGrid(CellRead[][] cells, OoxmlStylesResolver styles) {
        String[][] g = new String[cells.length][];
        for (int r = 0; r < cells.length; r++) {
            g[r] = new String[cells[r].length];
            for (int c = 0; c < cells[r].length; c++) {
                CellRead cr = cells[r][c];
                if (cr == null) {
                    continue;
                }
                String code = styles.formatCodeForCellStyle(cr.styleIndex());
                if (code != null && !code.isEmpty()) {
                    g[r][c] = sanitizePlainTsvField(code);
                }
            }
        }
        return g;
    }

    private static String[][] toStyleGrid(CellRead[][] cells) {
        String[][] g = new String[cells.length][];
        for (int r = 0; r < cells.length; r++) {
            g[r] = new String[cells[r].length];
            for (int c = 0; c < cells[r].length; c++) {
                CellRead cr = cells[r][c];
                if (cr == null) {
                    continue;
                }
                int idx = cr.styleIndex();
                g[r][c] = idx >= 0 ? Integer.toString(idx) : "0";
            }
        }
        return g;
    }

    /** TSV の列を壊さないよう、プレーンテキストセル内のタブ・改行を空白にする。 */
    private static String sanitizePlainTsvField(String s) {
        return s.replace('\t', ' ').replace('\r', ' ').replace('\n', ' ');
    }

    /**
     * {@code <c>} の子を {@code </c>} まで読み、格納値・型・数式テキストを返す。
     */
    private static CellRead readCellPayload(
            XMLStreamReader r, String cellTypeAttr, List<String> sharedStrings, int styleIndex)
            throws XMLStreamException {
        String vText = null;
        String inlineText = null;
        String formulaText = null;
        boolean hasFormula = false;
        int depth = 1;
        while (depth > 0 && r.hasNext()) {
            int ev = r.next();
            if (ev == XMLStreamConstants.START_ELEMENT) {
                String ln = r.getLocalName();
                if ("v".equals(ln)) {
                    vText = r.getElementText();
                } else if ("is".equals(ln)) {
                    inlineText = readInlineStringSi(r);
                } else if ("f".equals(ln)) {
                    hasFormula = true;
                    String ft = r.getElementText();
                    if (ft != null && !ft.trim().isEmpty()) {
                        formulaText = ft.trim();
                    }
                } else {
                    depth++;
                }
            } else if (ev == XMLStreamConstants.END_ELEMENT) {
                depth--;
            }
        }

        String valueResolved = resolveStoredValue(inlineText, vText, cellTypeAttr, sharedStrings);
        String typeLabel = classifyCellType(cellTypeAttr, hasFormula, vText, valueResolved);
        Integer ssIdx = parseSharedStringIndex(cellTypeAttr, vText);
        return new CellRead(valueResolved, typeLabel, formulaText, styleIndex, ssIdx);
    }

    /** {@code t="s"} かつ {@code <v>} が非負整数のときその値、それ以外 {@code null}。 */
    private static Integer parseSharedStringIndex(String cellTypeAttr, String vText) {
        if (!"s".equals(cellTypeAttr) || vText == null) {
            return null;
        }
        String v = vText.trim();
        if (v.isEmpty()) {
            return null;
        }
        try {
            int idx = Integer.parseInt(v);
            if (idx < 0) {
                return null;
            }
            return idx;
        } catch (NumberFormatException e) {
            return null;
        }
    }

    private static String resolveStoredValue(
            String inlineText, String vText, String cellTypeAttr, List<String> sharedStrings) {
        if (inlineText != null && !inlineText.isEmpty()) {
            return inlineText;
        }
        if (vText == null) {
            return null;
        }
        String v = vText.trim();
        if (v.isEmpty()) {
            return null;
        }
        if ("s".equals(cellTypeAttr)) {
            try {
                int idx = Integer.parseInt(v);
                if (idx >= 0 && idx < sharedStrings.size()) {
                    return sharedStrings.get(idx);
                }
                return "";
            } catch (NumberFormatException e) {
                return vText;
            }
        }
        if ("b".equals(cellTypeAttr)) {
            return "1".equals(v) ? "TRUE" : "FALSE";
        }
        return vText;
    }

    /**
     * {@code type.tsv} 用ラベル。空セルは {@code null}（TSV では空フィールド）。それ以外は {@code string},{@code number},{@code boolean},{@code error},{@code date} のいずれか。
     */
    private static String classifyCellType(String tAttr, boolean hasFormula, String vRaw, String valueResolved) {
        if (valueResolved != null && !valueResolved.isEmpty()) {
            if ("s".equals(tAttr) || "str".equals(tAttr)) {
                return "string";
            }
            if ("inlineStr".equals(tAttr)) {
                return "string";
            }
            if ("b".equals(tAttr)) {
                return "boolean";
            }
            if ("e".equals(tAttr)) {
                return "error";
            }
            if ("d".equals(tAttr)) {
                return "date";
            }
            if ("n".equals(tAttr)) {
                return "number";
            }
            // t 省略（既定 number）または数式キャッシュなど
            if (hasFormula || tAttr == null || tAttr.isEmpty()) {
                if (isNumericToken(valueResolved.trim())) {
                    return "number";
                }
                return "string";
            }
            return "string";
        }

        // 値が空
        if ("s".equals(tAttr) || "str".equals(tAttr) || "inlineStr".equals(tAttr)) {
            return null;
        }
        if ("b".equals(tAttr)) {
            return "boolean";
        }
        if ("e".equals(tAttr)) {
            return "error";
        }
        if (hasFormula) {
            return null;
        }
        if (vRaw != null && !vRaw.trim().isEmpty()) {
            // 型だけ残るケースは稀
            return isNumericToken(vRaw.trim()) ? "number" : "string";
        }
        return null;
    }

    private static boolean isNumericToken(String s) {
        if (s.isEmpty()) {
            return false;
        }
        try {
            Double.parseDouble(s);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    /** 現在位置が {@code <is>} の開始イベントにあるとき、{@code </is>} まで読み、{@code <t>} のテキストを連結する。 */
    private static String readInlineStringSi(XMLStreamReader r) throws XMLStreamException {
        StringBuilder sb = new StringBuilder();
        int depth = 1;
        while (depth > 0 && r.hasNext()) {
            int ev = r.next();
            if (ev == XMLStreamConstants.START_ELEMENT) {
                if ("t".equals(r.getLocalName())) {
                    sb.append(r.getElementText());
                } else {
                    depth++;
                }
            } else if (ev == XMLStreamConstants.END_ELEMENT) {
                depth--;
            }
        }
        return sb.toString();
    }

    /** {@code <c>} の直下の最初の子から {@code </c>} まで読み飛ばす。 */
    private static void consumeCellSubtree(XMLStreamReader r) throws XMLStreamException {
        int depth = 1;
        while (depth > 0 && r.hasNext()) {
            int ev = r.next();
            if (ev == XMLStreamConstants.START_ELEMENT) {
                depth++;
            } else if (ev == XMLStreamConstants.END_ELEMENT) {
                depth--;
            }
        }
    }

    private static void writeGridPlainTsv(Path outFile, String[][] grid) throws IOException {
        try (BufferedWriter w = Files.newBufferedWriter(outFile, StandardCharsets.UTF_8)) {
            for (int r = 0; r < grid.length; r++) {
                for (int c = 0; c < grid[r].length; c++) {
                    if (c > 0) {
                        w.write('\t');
                    }
                    String cell = grid[r][c];
                    if (cell != null && !cell.isEmpty()) {
                        w.write(cell);
                    }
                }
                w.write('\n');
            }
        }
    }

    private static void writeGridBase64Tsv(Path outFile, String[][] grid) throws IOException {
        Base64.Encoder enc = Base64.getEncoder();
        try (BufferedWriter w = Files.newBufferedWriter(outFile, StandardCharsets.UTF_8)) {
            for (int r = 0; r < grid.length; r++) {
                for (int c = 0; c < grid[r].length; c++) {
                    if (c > 0) {
                        w.write('\t');
                    }
                    String cell = grid[r][c];
                    if (cell != null && !cell.isEmpty()) {
                        w.write(enc.encodeToString(cell.getBytes(StandardCharsets.UTF_8)));
                    }
                }
                w.write('\n');
            }
        }
    }
}
