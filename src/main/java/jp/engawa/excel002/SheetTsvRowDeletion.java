package jp.engawa.excel002;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.NavigableSet;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * {@code data/} にあるシートのグリッド TSV から<strong>ワークシート上の指定行をまとめて削除</strong>する（1 始まり行番号）。
 *
 * <p>手順は「全 aspect を同じ順で読み込み → 削除行を除いて一時ファイルに書く → 行・列数の整合を確認 → 元ファイル削除とリネーム」。
 * commit より前で失敗した場合は作った一時ファイルだけを削除し、原本はそのまま残す。
 *
 * <p>{@code merge.tsv} の {@code (col,row)} は、削除により下に詰まった行に合わせて行番号を減らす。削除行上に結合代表 {@code 0,1} があれば拒否する。
 * {@code formula.tsv} のセル内文字列（A1 形式の参照など）は書き換えない。
 */
final class SheetTsvRowDeletion {
    private SheetTsvRowDeletion() {}

    private static final String[] GRID_ASPECTS = {
        "value", "type", "string", "formula", "format", "merge", "style", "formatted_value"
    };

    private static final Pattern MERGE_ORIGIN_PAREN =
            Pattern.compile("^\\(\\s*(\\d+)\\s*,\\s*(\\d+)\\s*\\)$");

    /** {@link WorksheetTsvRebuilder} と同一（{@link SheetsJsonExporter} の並び依存）。 */
    private static final Pattern SHEETS_JSON_SHEET_RECORD =
            Pattern.compile(
                    "\\{\"sheet_name\":\"([^\"]*)\",\"sheet_index\":(\\d+),\"max_row_num\":(\\d+),\"max_column_num\":(\\d+),\"sheet_id\":\"([^\"]+)\",\"relationship_id\":\"([^\"]+)\",\"worksheet_xml\":\"([^\"]+)\"\\}");

    static void deleteRows(Path dataDir, String sheetId, int... deletedRowsOneBased) throws IOException {
        if (deletedRowsOneBased == null || deletedRowsOneBased.length == 0) {
            throw new IOException("削除する行がありません（1 始まりの行番号を指定してください）");
        }

        NavigableSet<Integer> del = uniqueSorted(deletedRowsOneBased);

        Path sheetsJson = dataDir.resolve("sheets.json");
        if (!Files.isRegularFile(sheetsJson)) {
            throw new IOException("sheets.json がありません: " + sheetsJson);
        }
        ParsedSheet parsed = extractSheetFields(Files.readString(sheetsJson, StandardCharsets.UTF_8), sheetId);
        int oldRows = parsed.maxRowNum;
        int cols = parsed.maxColNum;

        validateDeleteSet(del, oldRows);

        validateNoMergeAnchorOnDeletedRows(dataDir, sheetId, del, cols, oldRows);

        int newRows = oldRows - del.size();

        Path sheetsJsonTmp = sheetsJson.resolveSibling("sheets.json.sheetrowdel.tmp");
        List<Path> aspectTmps = new ArrayList<>();

        try {
            for (String aspect : GRID_ASPECTS) {
                Path orig = dataDir.resolve(sheetId + "." + aspect + ".tsv");
                if (!Files.isRegularFile(orig)) {
                    throw new IOException("TSV がありません: " + orig);
                }
                Path tmp = orig.resolveSibling(orig.getFileName() + ".sheetrowdel.tmp");
                aspectTmps.add(tmp);
                List<String> lines = Files.readAllLines(orig, StandardCharsets.UTF_8);
                if (lines.size() < oldRows) {
                    throw new IOException("TSV の行数が sheets.json と一致しません: " + orig);
                }

                switch (aspect) {
                    case "merge" ->
                            writeMergedFiltered(lines, cols, oldRows, tmp, del);
                    default -> writePlainFiltered(lines, cols, oldRows, tmp, del);
                }
            }

            for (Path t : aspectTmps) {
                assertGridDimensions(t, newRows, cols);
            }

            String newJsonBody = rewrittenSheetsJson(Files.readString(sheetsJson, StandardCharsets.UTF_8), sheetId, newRows);
            Files.writeString(sheetsJsonTmp, newJsonBody, StandardCharsets.UTF_8);

            commitReplace(dataDir, sheetId);
            Files.deleteIfExists(sheetsJson);
            Files.move(sheetsJsonTmp, sheetsJson);
        } catch (IOException | RuntimeException e) {
            for (Path t : aspectTmps) {
                Files.deleteIfExists(t);
            }
            Files.deleteIfExists(sheetsJsonTmp);
            throw e;
        }
    }

    private record ParsedSheet(String sheetId, int maxRowNum, int maxColNum) {}

    private static ParsedSheet extractSheetFields(String json, String sheetId) throws IOException {
        Matcher m = SHEETS_JSON_SHEET_RECORD.matcher(json);
        while (m.find()) {
            if (sheetId.equals(m.group(5))) {
                return new ParsedSheet(
                        m.group(5),
                        Integer.parseInt(m.group(3)),
                        Integer.parseInt(m.group(4)));
            }
        }
        throw new IOException("sheets.json に sheet_id=\"" + sheetId + "\" のシートがありません");
    }

    private static String rewrittenSheetsJson(String json, String sheetId, int newRowCount) throws IOException {
        Matcher m = SHEETS_JSON_SHEET_RECORD.matcher(json);
        StringBuilder out = new StringBuilder(json.length() + 8);
        int copyFrom = 0;
        boolean replaced = false;
        while (m.find()) {
            out.append(json, copyFrom, m.start());
            if (sheetId.equals(m.group(5))) {
                replaced = true;
                out.append("{\"sheet_name\":\"")
                        .append(m.group(1))
                        .append("\",\"sheet_index\":")
                        .append(m.group(2))
                        .append(",\"max_row_num\":")
                        .append(newRowCount)
                        .append(",\"max_column_num\":")
                        .append(m.group(4))
                        .append(",\"sheet_id\":\"")
                        .append(m.group(5))
                        .append("\",\"relationship_id\":\"")
                        .append(m.group(6))
                        .append("\",\"worksheet_xml\":\"")
                        .append(m.group(7))
                        .append("\"}");
            } else {
                out.append(m.group());
            }
            copyFrom = m.end();
        }
        out.append(json, copyFrom, json.length());
        if (!replaced) {
            throw new IOException("sheets.json 内で sheet_id=\"" + sheetId + "\" を置換できませんでした");
        }
        return out.toString();
    }

    private static NavigableSet<Integer> uniqueSorted(int[] rows) throws IOException {
        TreeSet<Integer> s = new TreeSet<>();
        for (int ro : rows) {
            if (ro < 1) {
                throw new IOException("削除行は 1 始まりで指定してください: " + ro);
            }
            s.add(ro);
        }
        return s;
    }

    private static void validateDeleteSet(NavigableSet<Integer> del, int oldRows) throws IOException {
        if (del.first() < 1 || del.last() > oldRows) {
            throw new IOException(
                    String.format(
                            Locale.ROOT,
                            "削除行が範囲外です [1,%d]: %s",
                            oldRows,
                            del));
        }
        if (del.size() >= oldRows) {
            throw new IOException("少なくとも 1 行は残してください（すべての行が削除対象になりました）");
        }
    }

    private static void validateNoMergeAnchorOnDeletedRows(
            Path dataDir, String sheetId, NavigableSet<Integer> del, int cols, int oldRows) throws IOException {
        Path merge = dataDir.resolve(sheetId + ".merge.tsv");
        if (!Files.isRegularFile(merge)) {
            return;
        }
        List<String> lines = Files.readAllLines(merge, StandardCharsets.UTF_8);
        for (int r = 1; r <= oldRows && r <= lines.size(); r++) {
            if (!del.contains(r)) {
                continue;
            }
            String[] cells = splitTsvPreservingEnds(lines.get(r - 1), cols);
            for (String cell : cells) {
                if (cell != null && "0,1".equals(cell.trim())) {
                    throw new IOException(
                            String.format(Locale.ROOT, "削除不能: 結合セルの代表セル (%d 行目) が merge.tsv にあります", r));
                }
            }
        }
    }

    private static void writePlainFiltered(
            List<String> lines, int cols, int oldRows, Path tmp, NavigableSet<Integer> del) throws IOException {
        ArrayList<String> out = new ArrayList<>(oldRows - del.size());
        for (int oldR = 1; oldR <= oldRows; oldR++) {
            if (del.contains(oldR)) {
                continue;
            }
            String line = lines.get(oldR - 1);
            enforceColumnCount(line, cols, tmp, oldR);
            out.add(line);
        }
        Files.write(tmp, out, StandardCharsets.UTF_8);
    }

    private static void writeMergedFiltered(
            List<String> lines, int cols, int oldRows, Path tmp, NavigableSet<Integer> del) throws IOException {
        ArrayList<String> out = new ArrayList<>(oldRows - del.size());
        for (int oldR = 1; oldR <= oldRows; oldR++) {
            if (del.contains(oldR)) {
                continue;
            }
            String line = lines.get(oldR - 1);
            enforceColumnCount(line, cols, tmp, oldR);
            out.add(remapMergeValuesInLine(line, cols, del));
        }
        Files.write(tmp, out, StandardCharsets.UTF_8);
    }

    private static void enforceColumnCount(String line, int expectCols, Path contextFile, int oneBasedOldRowIndex)
            throws IOException {
        int n = splitTsvNoLimit(line).length;
        if (n != expectCols) {
            throw new IOException(
                    "TSV の列数が sheets.json と一致しません: "
                            + contextFile
                            + " （元の行 "
                            + oneBasedOldRowIndex
                            + "、列が "
                            + n
                            + " で期待 "
                            + expectCols
                            + ")");
        }
    }

    private static String remapMergeValuesInLine(String line, int cols, NavigableSet<Integer> deletedOneBased)
            throws IOException {
        String[] parts = splitTsvPreservingEnds(line, cols);
        StringBuilder rebuilt = new StringBuilder(line.length() + cols);
        for (int ci = 0; ci < cols; ci++) {
            if (ci > 0) {
                rebuilt.append('\t');
            }
            rebuilt.append(remapMergeField(parts[ci], deletedOneBased));
        }
        return rebuilt.toString();
    }

    private static String remapMergeField(String rawCell, NavigableSet<Integer> deletedOneBased) throws IOException {
        if (rawCell == null) {
            return "";
        }
        String t = rawCell.trim();
        Matcher mp = MERGE_ORIGIN_PAREN.matcher(t);
        if (!mp.matches()) {
            return rawCell;
        }
        String cStr = mp.group(1);
        String rStr = mp.group(2);
        int anchorRowOld = Integer.parseInt(rStr);
        if (deletedOneBased.contains(anchorRowOld)) {
            throw new IOException(
                    "merge.tsv: 削除行にぶら下がる結合元を指すセルがあります（行 " + anchorRowOld + " の参照 \"" + rawCell.trim() + "\"）");
        }
        int stripped = anchorRowOld - deletedOneBased.headSet(anchorRowOld).size();
        return "(" + cStr + "," + stripped + ")";
    }

    private static String[] splitTsvPreservingEnds(String line, int expectCols) {
        String[] parts = splitTsvNoLimit(line);
        String[] padded = new String[expectCols];
        for (int i = 0; i < expectCols; i++) {
            padded[i] = i < parts.length ? parts[i] : "";
        }
        return padded;
    }

    private static String[] splitTsvNoLimit(String line) {
        if (line == null) {
            return new String[0];
        }
        return line.split("\t", -1);
    }

    private static void assertGridDimensions(Path tsvPath, int expectRows, int expectCols) throws IOException {
        List<String> lines = Files.readAllLines(tsvPath, StandardCharsets.UTF_8);
        if (lines.size() != expectRows) {
            throw new IOException(
                    "行削除後に全 aspect で行数が一致しません: " + tsvPath + " は " + lines.size() + " 行ですが期待は " + expectRows);
        }
        for (int r = 1; r <= lines.size(); r++) {
            int nFields = splitTsvNoLimit(lines.get(r - 1)).length;
            if (nFields != expectCols) {
                throw new IOException(
                        "列数が不一致です: "
                                + tsvPath
                                + " （行 "
                                + r
                                + " が "
                                + nFields
                                + " 列、期待 "
                                + expectCols
                                + "）。commit 中止のため削除します（一時ファイル）。");
            }
        }
    }

    private static void commitReplace(Path dataDir, String sheetId) throws IOException {
        for (String aspect : GRID_ASPECTS) {
            Path orig = dataDir.resolve(sheetId + "." + aspect + ".tsv");
            Path tmp = orig.resolveSibling(orig.getFileName() + ".sheetrowdel.tmp");
            Files.deleteIfExists(orig);
            Files.move(tmp, orig);
        }
    }
}
