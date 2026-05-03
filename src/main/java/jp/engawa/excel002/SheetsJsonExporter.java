package jp.engawa.excel002;

import java.io.BufferedWriter;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
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
 * 展開済みブック（{@code tmp.dir/md5/excel}）からメタデータを読み、各シートのワークシート XML への excel ルート相対パス（例: {@code xl/worksheets/sheet1.xml}）とともに {@code data/sheets.json} を書く。
 */
final class SheetsJsonExporter {
    private static final String NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static final String NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static final String NS_REL_OD = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static final Pattern DIMENSION_REF =
            Pattern.compile("<dimension[^>]*ref=\"([^\"]*)\"", Pattern.CASE_INSENSITIVE);
    private static final Pattern CELL_R_ATTR =
            Pattern.compile("<c[^>]*\\br=\"([A-Za-z]+)(\\d+)\"");

    private SheetsJsonExporter() {}

    static void write(Path excelRoot, Path sheetsJsonFile, File originalFile, String md5Hex)
            throws IOException {
        Path xl = excelRoot.resolve("xl");
        if (!Files.isDirectory(xl)) {
            throw new IOException("xl ディレクトリがありません: " + excelRoot);
        }

        Path workbookXml = xl.resolve("workbook.xml");
        Path workbookRels = xl.resolve("_rels/workbook.xml.rels");
        Path excelRootNorm = excelRoot.toAbsolutePath().normalize();
        Map<String, String> worksheetTargets = readWorksheetTargets(workbookRels);
        List<SheetEntry> sheets = readSheetEntries(workbookXml, worksheetTargets, xl, excelRootNorm);

        String fileName = originalFile.getName();
        long lastMod = originalFile.lastModified();

        Path parent = sheetsJsonFile.toAbsolutePath().normalize().getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        try (BufferedWriter w =
                Files.newBufferedWriter(sheetsJsonFile, StandardCharsets.UTF_8)) {
            w.write('{');
            w.write("\"file_name\":");
            writeJsonString(w, fileName);
            w.write(',');
            w.write("\"last_modified\":");
            w.write(Long.toString(lastMod));
            w.write(',');
            w.write("\"file_md5\":");
            writeJsonString(w, md5Hex);
            w.write(',');
            w.write("\"sheets_size\":");
            w.write(Integer.toString(sheets.size()));
            w.write(',');
            w.write("\"sheets\":[");
            for (int i = 0; i < sheets.size(); i++) {
                if (i > 0) {
                    w.write(',');
                }
                SheetEntry s = sheets.get(i);
                Path worksheetFile =
                        excelRootNorm.resolve(s.worksheetRelativePath()).normalize();
                if (!worksheetFile.startsWith(excelRootNorm)) {
                    throw new IOException("ワークシートパスが excel ルートの外を指しています: " + worksheetFile);
                }
                IntPair max = dimensionMaxFromWorksheetFile(worksheetFile);
                w.write('{');
                w.write("\"sheet_name\":");
                writeJsonString(w, s.name());
                w.write(',');
                w.write("\"sheet_index\":");
                w.write(Integer.toString(i));
                w.write(',');
                w.write("\"max_row_num\":");
                w.write(Integer.toString(max.row()));
                w.write(',');
                w.write("\"max_column_num\":");
                w.write(Integer.toString(max.col()));
                w.write(',');
                w.write("\"sheet_id\":");
                writeJsonString(w, s.sheetId());
                w.write(',');
                w.write("\"relationship_id\":");
                writeJsonString(w, s.relationshipId());
                w.write(',');
                w.write("\"worksheet_xml\":");
                writeJsonString(w, s.worksheetRelativePath());
                w.write('}');
            }
            w.write(']');
            w.write('}');
        }
    }

    /** {@code worksheetRelativePath} は excel 展開ルートからの相対パス（例: {@code xl/worksheets/sheet1.xml}）。JSON キーは {@code worksheet_xml} のまま。 */
    private record SheetEntry(
            String name,
            String sheetId,
            String relationshipId,
            String worksheetRelativePath) {}

    private record IntPair(int row, int col) {}

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

    private static List<SheetEntry> readSheetEntries(
            Path workbookXml,
            Map<String, String> worksheetTargets,
            Path xlDir,
            Path excelRootNorm)
            throws IOException {
        if (!Files.isRegularFile(workbookXml)) {
            throw new IOException("workbook.xml がありません: " + workbookXml);
        }
        List<SheetEntry> list = new ArrayList<>();
        XMLInputFactory f = XMLInputFactory.newInstance();
        try (InputStream in = Files.newInputStream(workbookXml)) {
            XMLStreamReader r = f.createXMLStreamReader(in);
            while (r.hasNext()) {
                if (r.next() == XMLStreamConstants.START_ELEMENT
                        && "sheet".equals(r.getLocalName())
                        && NS_MAIN.equals(r.getNamespaceURI())) {
                    String name = r.getAttributeValue(null, "name");
                    String sheetId = r.getAttributeValue(null, "sheetId");
                    String rid = r.getAttributeValue(NS_REL_OD, "id");
                    if (name == null || sheetId == null || rid == null) {
                        continue;
                    }
                    String target = worksheetTargets.get(rid);
                    if (target == null) {
                        continue;
                    }
                    Path sheetPath = xlDir.resolve(target).normalize().toAbsolutePath().normalize();
                    Path xlNorm = xlDir.toAbsolutePath().normalize();
                    if (!sheetPath.startsWith(xlNorm)) {
                        throw new IOException("シートパスが xl の外を指しています: " + target);
                    }
                    if (!sheetPath.startsWith(excelRootNorm)) {
                        throw new IOException("シートパスが展開ルートの外を指しています: " + sheetPath);
                    }
                    if (!Files.isRegularFile(sheetPath)) {
                        throw new IOException("ワークシート XML が見つかりません: " + sheetPath);
                    }
                    Path relToExcel = excelRootNorm.relativize(sheetPath);
                    if (relToExcel.isAbsolute()) {
                        throw new IOException("ワークシートパスを excel ルートからの相対パスにできません: " + sheetPath);
                    }
                    String pathJson = relToExcel.toString().replace('\\', '/');
                    list.add(new SheetEntry(name, sheetId, rid, pathJson));
                }
            }
            r.close();
        } catch (XMLStreamException e) {
            throw new IOException("workbook.xml の解析に失敗しました", e);
        }
        return list;
    }

    /** 1 始まりの最大行・最大列。データが無いときは 0。 */
    private static IntPair dimensionMaxFromWorksheetFile(Path worksheetXmlFile) throws IOException {
        String worksheetXml = Files.readString(worksheetXmlFile, StandardCharsets.UTF_8);
        Matcher dm = DIMENSION_REF.matcher(worksheetXml);
        if (dm.find()) {
            String ref = dm.group(1).trim();
            if (!ref.isEmpty()) {
                return maxBoundsFromRef(ref);
            }
        }
        int maxRow = 0;
        int maxCol = 0;
        Matcher cm = CELL_R_ATTR.matcher(worksheetXml);
        while (cm.find()) {
            int row = Integer.parseInt(cm.group(2));
            int col = columnLettersToOneBased(cm.group(1));
            maxRow = Math.max(maxRow, row);
            maxCol = Math.max(maxCol, col);
        }
        return new IntPair(maxRow, maxCol);
    }

    private static IntPair maxBoundsFromRef(String ref) {
        String[] parts = ref.split(":");
        if (parts.length == 1) {
            return cellRefOneBased(parts[0]);
        }
        IntPair a = cellRefOneBased(parts[0]);
        IntPair b = cellRefOneBased(parts[1]);
        return new IntPair(Math.max(a.row(), b.row()), Math.max(a.col(), b.col()));
    }

    private static IntPair cellRefOneBased(String cell) {
        int i = 0;
        while (i < cell.length() && Character.isLetter(cell.charAt(i))) {
            i++;
        }
        if (i == 0 || i >= cell.length()) {
            return new IntPair(0, 0);
        }
        String letters = cell.substring(0, i).toUpperCase();
        int row = Integer.parseInt(cell.substring(i));
        int col = columnLettersToOneBased(letters);
        return new IntPair(row, col);
    }

    private static int columnLettersToOneBased(String letters) {
        int col = 0;
        for (int k = 0; k < letters.length(); k++) {
            col = col * 26 + (letters.charAt(k) - 'A' + 1);
        }
        return col;
    }

    private static void writeJsonString(Appendable w, String s) throws IOException {
        w.append('"');
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            switch (c) {
                case '\\':
                    w.append("\\\\");
                    break;
                case '"':
                    w.append("\\\"");
                    break;
                case '\b':
                    w.append("\\b");
                    break;
                case '\f':
                    w.append("\\f");
                    break;
                case '\n':
                    w.append("\\n");
                    break;
                case '\r':
                    w.append("\\r");
                    break;
                case '\t':
                    w.append("\\t");
                    break;
                default:
                    if (c < 0x20) {
                        w.append("\\u");
                        w.append(String.format("%04x", (int) c));
                    } else {
                        w.append(c);
                    }
                    break;
            }
        }
        w.append('"');
    }
}
