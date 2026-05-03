package jp.engawa.excel002;

import java.io.BufferedWriter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

/**
 * {@code data/} の TSV 群と {@code sheets.json} からワークシート XML（例: {@code xl/worksheets/sheet1.xml}）を再構築する。
 *
 * <p>入力: {@code {sheet_id}.value.tsv}（Base64）、{@code .type.tsv}、{@code .string.tsv}、{@code .formula.tsv}、
 * {@code .merge.tsv}、{@code .style.tsv}。グリッド寸法は {@code sheets.json} の {@code max_row_num} /
 * {@code max_column_num} と各 TSV の行・列数が一致すること。
 */
public final class WorksheetTsvRebuilder {
    private static final String NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static final Pattern SHEETS_JSON_SHEET_OBJECT =
            Pattern.compile(
                    "\\{\"sheet_name\":\"([^\"]*)\",\"sheet_index\":(\\d+),\"max_row_num\":(\\d+),\"max_column_num\":(\\d+),\"sheet_id\":\"([^\"]+)\",\"relationship_id\":\"([^\"]+)\",\"worksheet_xml\":\"([^\"]+)\"\\}");

    private WorksheetTsvRebuilder() {}

    /**
     * {@code excelRoot} 上のワークシート XML を、{@code dataDir} の TSV と {@code sheets.json} から再書き込みする。
     *
     * @param excelRoot 展開済みブックのルート（{@code xl/} を含む）
     * @param dataDir {@code sheets.json} および {@code {sheet_id}.*.tsv} を含むディレクトリ（通常は {@code ../data}）
     * @param sheetId {@code sheets.json} の {@code sheet_id}（例: {@code "1"}）
     */
    public static void rebuildWorksheet(Path excelRoot, Path dataDir, String sheetId) throws IOException {
        Path sheetsJson = dataDir.resolve("sheets.json");
        if (!Files.isRegularFile(sheetsJson)) {
            throw new IOException("sheets.json がありません: " + sheetsJson);
        }
        SheetJsonEntry entry = findSheetEntry(sheetsJson, sheetId);
        Path worksheetPath = excelRoot.resolve(entry.worksheetXml()).normalize();
        Path excelNorm = excelRoot.toAbsolutePath().normalize();
        if (!worksheetPath.startsWith(excelNorm)) {
            throw new IOException("ワークシートパスが excel ルートの外です: " + worksheetPath);
        }
        Files.createDirectories(worksheetPath.getParent());

        String[][] value = readBase64Grid(dataDir.resolve(sheetId + ".value.tsv"), entry.rows(), entry.cols());
        String[][] type = readPlainGrid(dataDir.resolve(sheetId + ".type.tsv"), entry.rows(), entry.cols());
        String[][] stringId = readPlainGrid(dataDir.resolve(sheetId + ".string.tsv"), entry.rows(), entry.cols());
        String[][] formula = readPlainGrid(dataDir.resolve(sheetId + ".formula.tsv"), entry.rows(), entry.cols());
        String[][] merge = readPlainGrid(dataDir.resolve(sheetId + ".merge.tsv"), entry.rows(), entry.cols());
        String[][] style = readPlainGrid(dataDir.resolve(sheetId + ".style.tsv"), entry.rows(), entry.cols());

        List<MergeRegion> merges = findMergeRegions(merge);
        writeWorksheetXml(worksheetPath, entry.rows(), entry.cols(), value, type, stringId, formula, merge, style, merges);
    }

    private record SheetJsonEntry(String sheetId, String worksheetXml, int rows, int cols) {}

    private static SheetJsonEntry findSheetEntry(Path sheetsJson, String sheetId) throws IOException {
        String json = Files.readString(sheetsJson, StandardCharsets.UTF_8);
        Matcher m = SHEETS_JSON_SHEET_OBJECT.matcher(json);
        while (m.find()) {
            String id = m.group(5);
            if (sheetId.equals(id)) {
                int rows = Integer.parseInt(m.group(3));
                int cols = Integer.parseInt(m.group(4));
                String path = m.group(7);
                return new SheetJsonEntry(id, path, rows, cols);
            }
        }
        throw new IOException("sheets.json に sheet_id=\"" + sheetId + "\" がありません: " + sheetsJson);
    }

    private static String[][] readPlainGrid(Path file, int expectRows, int expectCols) throws IOException {
        if (!Files.isRegularFile(file)) {
            throw new IOException("TSV がありません: " + file);
        }
        List<String> lines = Files.readAllLines(file, StandardCharsets.UTF_8);
        if (lines.size() < expectRows) {
            throw new IOException("TSV の行数が足りません: " + file + " (" + lines.size() + " < " + expectRows + ")");
        }
        String[][] g = new String[expectRows][expectCols];
        for (int r = 0; r < expectRows; r++) {
            String[] parts = lines.get(r).split("\t", -1);
            if (parts.length < expectCols) {
                throw new IOException("TSV の列数が足りません: " + file + " 行 " + (r + 1));
            }
            System.arraycopy(parts, 0, g[r], 0, expectCols);
        }
        return g;
    }

    private static String[][] readBase64Grid(Path file, int expectRows, int expectCols) throws IOException {
        String[][] raw = readPlainGrid(file, expectRows, expectCols);
        Base64.Decoder dec = Base64.getDecoder();
        String[][] out = new String[expectRows][expectCols];
        for (int r = 0; r < expectRows; r++) {
            for (int c = 0; c < expectCols; c++) {
                String field = raw[r][c];
                if (field == null || field.isEmpty()) {
                    out[r][c] = "";
                } else {
                    try {
                        out[r][c] = new String(dec.decode(field.trim()), StandardCharsets.UTF_8);
                    } catch (IllegalArgumentException e) {
                        throw new IOException("Base64 のデコードに失敗しました: " + file + " (" + r + "," + c + ")", e);
                    }
                }
            }
        }
        return out;
    }

    private record MergeRegion(int row1OneBased, int col1OneBased, int row2OneBased, int col2OneBased) {}

    private static List<MergeRegion> findMergeRegions(String[][] merge) {
        if (merge.length == 0 || merge[0].length == 0) {
            return List.of();
        }
        int rows = merge.length;
        int cols = merge[0].length;
        boolean[][] used = new boolean[rows][cols];
        List<MergeRegion> list = new ArrayList<>();
        for (int r = 0; r < rows; r++) {
            for (int c = 0; c < cols; c++) {
                if (used[r][c]) {
                    continue;
                }
                String m = merge[r][c];
                if (m == null || m.isEmpty() || !"0,1".equals(m.trim())) {
                    continue;
                }
                int width = 1;
                while (c + width < cols) {
                    String right = merge[r][c + width];
                    if (right != null && "-1,-1".equals(right.trim())) {
                        width++;
                    } else {
                        break;
                    }
                }
                int height = 1;
                outer:
                while (r + height < rows) {
                    for (int dc = 0; dc < width; dc++) {
                        String below = merge[r + height][c + dc];
                        if (below == null || !"-1,-1".equals(below.trim())) {
                            break outer;
                        }
                    }
                    height++;
                }
                for (int rr = r; rr < r + height; rr++) {
                    for (int cc = c; cc < c + width; cc++) {
                        used[rr][cc] = true;
                    }
                }
                list.add(new MergeRegion(r + 1, c + 1, r + height, c + width));
            }
        }
        return list;
    }

    private static void writeWorksheetXml(
            Path outFile,
            int rowCount,
            int colCount,
            String[][] value,
            String[][] type,
            String[][] stringId,
            String[][] formula,
            String[][] merge,
            String[][] style,
            List<MergeRegion> merges)
            throws IOException {
        String dimRef = "A1:" + columnLetters(colCount) + rowCount;
        XMLOutputFactory xf = XMLOutputFactory.newInstance();
        try (BufferedWriter bw = Files.newBufferedWriter(outFile, StandardCharsets.UTF_8)) {
            XMLStreamWriter w = xf.createXMLStreamWriter(bw);
            w.writeStartDocument(StandardCharsets.UTF_8.name(), "1.0");
            w.setDefaultNamespace(NS_MAIN);
            w.writeStartElement(NS_MAIN, "worksheet");
            w.writeDefaultNamespace(NS_MAIN);
            w.writeStartElement(NS_MAIN, "dimension");
            w.writeAttribute("ref", dimRef);
            w.writeEndElement();
            w.writeEmptyElement(NS_MAIN, "sheetFormatPr");
            w.writeAttribute("baseColWidth", "10");
            w.writeAttribute("defaultRowHeight", "20");
            w.writeStartElement(NS_MAIN, "sheetData");
            for (int r = 0; r < rowCount; r++) {
                int row1 = r + 1;
                RowSpan span = rowEmitSpan(colCount, r, value, type, stringId, formula, merge, style);
                if (span == null) {
                    continue;
                }
                w.writeStartElement(NS_MAIN, "row");
                w.writeAttribute("r", Integer.toString(row1));
                w.writeAttribute("spans", span.minCol() + ":" + span.maxCol());
                for (int c = span.minColZero(); c <= span.maxColZero(); c++) {
                    if (!shouldEmitCell(r, c, value, type, stringId, formula, merge, style)) {
                        continue;
                    }
                    writeCellElement(w, row1, c + 1, value[r][c], type[r][c], stringId[r][c], formula[r][c], merge[r][c], style[r][c]);
                }
                w.writeEndElement();
            }
            w.writeEndElement();
            if (!merges.isEmpty()) {
                w.writeStartElement(NS_MAIN, "mergeCells");
                w.writeAttribute("count", Integer.toString(merges.size()));
                for (MergeRegion m : merges) {
                    String ref =
                            cellRef(m.col1OneBased(), m.row1OneBased())
                                    + ":"
                                    + cellRef(m.col2OneBased(), m.row2OneBased());
                    w.writeStartElement(NS_MAIN, "mergeCell");
                    w.writeAttribute("ref", ref);
                    w.writeEndElement();
                }
                w.writeEndElement();
            }
            w.writeEmptyElement(NS_MAIN, "pageMargins");
            w.writeAttribute("left", "0.7");
            w.writeAttribute("right", "0.7");
            w.writeAttribute("top", "0.75");
            w.writeAttribute("bottom", "0.75");
            w.writeAttribute("header", "0.3");
            w.writeAttribute("footer", "0.3");
            w.writeEndElement();
            w.writeEndDocument();
            w.close();
        } catch (XMLStreamException e) {
            throw new IOException("ワークシート XML の書き込みに失敗しました: " + outFile, e);
        }
    }

    private record RowSpan(int minCol, int maxCol) {
        int minColZero() {
            return minCol - 1;
        }

        int maxColZero() {
            return maxCol - 1;
        }
    }

    private static RowSpan rowEmitSpan(
            int colCount,
            int rowZero,
            String[][] value,
            String[][] type,
            String[][] stringId,
            String[][] formula,
            String[][] merge,
            String[][] style) {
        int minCol = Integer.MAX_VALUE;
        int maxCol = Integer.MIN_VALUE;
        for (int c = 0; c < colCount; c++) {
            if (shouldEmitCell(rowZero, c, value, type, stringId, formula, merge, style)) {
                minCol = Math.min(minCol, c + 1);
                maxCol = Math.max(maxCol, c + 1);
            }
        }
        if (minCol == Integer.MAX_VALUE) {
            return null;
        }
        return new RowSpan(minCol, maxCol);
    }

    private static boolean shouldEmitCell(
            int r,
            int c,
            String[][] value,
            String[][] type,
            String[][] stringId,
            String[][] formula,
            String[][] merge,
            String[][] style) {
        String mg = merge[r][c];
        if (mg != null) {
            String mt = mg.trim();
            if ("0,1".equals(mt)) {
                return true;
            }
            if ("-1,-1".equals(mt)) {
                int si = styleIndex(style[r][c]);
                return si > 0;
            }
        }
        String f = formula[r][c];
        if (f != null && !f.trim().isEmpty()) {
            return true;
        }
        String sid = stringId[r][c];
        if (sid != null && !sid.trim().isEmpty()) {
            return true;
        }
        String v = value[r][c];
        if (v != null && !v.isEmpty()) {
            return true;
        }
        String t = type[r][c];
        if (hasMeaningfulTypeWithoutValue(t)) {
            return true;
        }
        int si = styleIndex(style[r][c]);
        return si > 0;
    }

    private static boolean hasMeaningfulTypeWithoutValue(String t) {
        if (t == null || t.isBlank()) {
            return false;
        }
        String x = t.trim();
        return "boolean".equals(x) || "error".equals(x);
    }

    private static int styleIndex(String styleField) {
        if (styleField == null || styleField.isBlank()) {
            return -1;
        }
        try {
            return Integer.parseInt(styleField.trim());
        } catch (NumberFormatException e) {
            return -1;
        }
    }

    private static void writeCellElement(
            XMLStreamWriter w,
            int rowOneBased,
            int colOneBased,
            String valueText,
            String typeLabel,
            String stringIndexField,
            String formulaText,
            String mergeField,
            String styleField)
            throws XMLStreamException {
        String ref = cellRef(colOneBased, rowOneBased);
        String mg = mergeField != null ? mergeField.trim() : "";
        int si = styleIndex(styleField);

        String f = formulaText != null ? formulaText.trim() : "";
        boolean hasFormula = !f.isEmpty();
        String sid = stringIndexField != null ? stringIndexField.trim() : "";
        boolean hasSharedString = !sid.isEmpty();
        String tlab = typeLabel != null ? typeLabel.trim() : "";

        if ("-1,-1".equals(mg)) {
            w.writeStartElement(NS_MAIN, "c");
            w.writeAttribute("r", ref);
            if (si > 0) {
                w.writeAttribute("s", Integer.toString(si));
            }
            w.writeEndElement();
            return;
        }

        String tAttr = null;
        boolean inlineStr = false;
        if (hasSharedString && !hasFormula) {
            tAttr = "s";
        } else if ("string".equals(tlab)
                && !hasSharedString
                && valueText != null
                && !valueText.isEmpty()
                && !hasFormula) {
            tAttr = "inlineStr";
            inlineStr = true;
        } else if ("boolean".equals(tlab)) {
            tAttr = "b";
        } else if ("error".equals(tlab)) {
            tAttr = "e";
        }

        w.writeStartElement(NS_MAIN, "c");
        w.writeAttribute("r", ref);
        if (si > 0) {
            w.writeAttribute("s", Integer.toString(si));
        }
        if (tAttr != null) {
            w.writeAttribute("t", tAttr);
        }

        if (hasFormula) {
            w.writeStartElement(NS_MAIN, "f");
            w.writeCharacters(f);
            w.writeEndElement();
        }

        if ("s".equals(tAttr)) {
            w.writeStartElement(NS_MAIN, "v");
            w.writeCharacters(sid);
            w.writeEndElement();
        } else if (inlineStr) {
            w.writeStartElement(NS_MAIN, "is");
            w.writeStartElement(NS_MAIN, "t");
            w.writeCharacters(valueText);
            w.writeEndElement();
            w.writeEndElement();
        } else if ("b".equals(tAttr)) {
            w.writeStartElement(NS_MAIN, "v");
            w.writeCharacters(booleanCellV(valueText));
            w.writeEndElement();
        } else if ("e".equals(tAttr)) {
            w.writeStartElement(NS_MAIN, "v");
            if (valueText != null && !valueText.isEmpty()) {
                w.writeCharacters(valueText);
            }
            w.writeEndElement();
        } else if (valueText != null && !valueText.isEmpty()) {
            w.writeStartElement(NS_MAIN, "v");
            w.writeCharacters(valueText);
            w.writeEndElement();
        }

        w.writeEndElement();
    }

    private static String booleanCellV(String valueText) {
        if (valueText == null) {
            return "0";
        }
        String v = valueText.trim();
        if ("1".equals(v) || "TRUE".equalsIgnoreCase(v)) {
            return "1";
        }
        return "0";
    }

    private static String cellRef(int colOneBased, int rowOneBased) {
        return columnLetters(colOneBased) + rowOneBased;
    }

    static String columnLetters(int colOneBased) {
        if (colOneBased <= 0) {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        int n = colOneBased;
        while (n > 0) {
            n--;
            sb.insert(0, (char) ('A' + (n % 26)));
            n /= 26;
        }
        return sb.toString();
    }
}
