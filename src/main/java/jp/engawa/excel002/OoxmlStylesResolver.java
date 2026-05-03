package jp.engawa.excel002;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

/**
 * {@code xl/styles.xml} から {@code cellXfs} と {@code numFmt} を読み、セルの {@code s} インデックスから {@code numFmtId}
 * と書式コード文字列を引けるようにする（簡易・基本用途）。
 */
final class OoxmlStylesResolver {
    private static final String NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /** Excel 組み込み {@code numFmtId} → 書式コード（抜粋。無い ID は {@code General}）。 */
    private static final Map<Integer, String> BUILTIN_NUMFMT = builtinNumFmts();

    private final List<Integer> cellXfNumFmtIds = new ArrayList<>();
    private final Map<Integer, String> numFmtIdToCode = new HashMap<>();

    private OoxmlStylesResolver() {}

    static OoxmlStylesResolver load(Path stylesXml) throws IOException {
        OoxmlStylesResolver r = new OoxmlStylesResolver();
        r.numFmtIdToCode.putAll(BUILTIN_NUMFMT);
        if (!Files.isRegularFile(stylesXml)) {
            return r;
        }
        XMLInputFactory f = XMLInputFactory.newInstance();
        try (InputStream in = Files.newInputStream(stylesXml)) {
            XMLStreamReader xr = f.createXMLStreamReader(in);
            boolean inCellXfs = false;
            while (xr.hasNext()) {
                int ev = xr.next();
                if (ev == XMLStreamConstants.START_ELEMENT
                        && NS.equals(xr.getNamespaceURI())) {
                    String ln = xr.getLocalName();
                    if ("cellXfs".equals(ln)) {
                        inCellXfs = true;
                    } else if ("numFmt".equals(ln)) {
                        String idStr = xr.getAttributeValue(null, "numFmtId");
                        String code = xr.getAttributeValue(null, "formatCode");
                        if (idStr != null && code != null && !code.isEmpty()) {
                            try {
                                r.numFmtIdToCode.put(Integer.parseInt(idStr.trim()), code);
                            } catch (NumberFormatException ignored) {
                                // skip
                            }
                        }
                    } else if ("xf".equals(ln) && inCellXfs) {
                        String n = xr.getAttributeValue(null, "numFmtId");
                        int id = 0;
                        if (n != null && !n.isBlank()) {
                            try {
                                id = Integer.parseInt(n.trim());
                            } catch (NumberFormatException ignored) {
                                id = 0;
                            }
                        }
                        r.cellXfNumFmtIds.add(id);
                    }
                } else if (ev == XMLStreamConstants.END_ELEMENT
                        && NS.equals(xr.getNamespaceURI())
                        && "cellXfs".equals(xr.getLocalName())) {
                    inCellXfs = false;
                }
            }
            xr.close();
        } catch (XMLStreamException e) {
            throw new IOException("styles.xml の解析に失敗しました", e);
        }
        return r;
    }

    /** {@code <c s="…">} のインデックス。無効時は 0 とみなす（先頭の xf）。 */
    int numFmtIdForCellStyle(int cellStyleIndex) {
        if (cellStyleIndex < 0 || cellStyleIndex >= cellXfNumFmtIds.size()) {
            return cellXfNumFmtIds.isEmpty() ? 0 : cellXfNumFmtIds.get(0);
        }
        return cellXfNumFmtIds.get(cellStyleIndex);
    }

    String formatCodeForCellStyle(int cellStyleIndex) {
        int id = numFmtIdForCellStyle(cellStyleIndex);
        return numFmtIdToCode.getOrDefault(id, "General");
    }

    private static Map<Integer, String> builtinNumFmts() {
        Map<Integer, String> m = new HashMap<>();
        m.put(0, "General");
        m.put(1, "0");
        m.put(2, "0.00");
        m.put(3, "#,##0");
        m.put(4, "#,##0.00");
        m.put(9, "0%");
        m.put(10, "0.00E+00");
        m.put(11, "0.00E+00");
        m.put(12, "# ?/?");
        m.put(13, "# ??/??");
        m.put(14, "m/d/yyyy");
        m.put(15, "d-mmm-yy");
        m.put(16, "d-mmm");
        m.put(17, "mmm-yy");
        m.put(18, "h:mm AM/PM");
        m.put(19, "h:mm:ss AM/PM");
        m.put(20, "h:mm");
        m.put(21, "h:mm:ss");
        m.put(22, "m/d/yyyy h:mm");
        for (int i = 37; i <= 44; i++) {
            if (!m.containsKey(i)) {
                m.put(i, "General");
            }
        }
        m.put(37, "#,##0_);(#,##0)");
        m.put(38, "#,##0_);[Red](#,##0)");
        m.put(39, "#,##0.00_);(#,##0.00)");
        m.put(40, "#,##0.00_);[Red](#,##0.00)");
        m.put(41, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)");
        m.put(42, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)");
        m.put(43, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)");
        m.put(44, "General");
        m.put(45, "mm:ss");
        m.put(46, "[h]:mm:ss");
        m.put(47, "mmss.0");
        m.put(48, "##0.0E+0");
        m.put(49, "@");
        return m;
    }

    /** 書式コードが日付っぽいか（ごく簡易）。 */
    static boolean looksLikeDateFormat(int numFmtId, String formatCode) {
        if (numFmtId >= 14 && numFmtId <= 22) {
            return true;
        }
        if (numFmtId >= 27 && numFmtId <= 36) {
            return true;
        }
        if (numFmtId >= 45 && numFmtId <= 47) {
            return true;
        }
        if (numFmtId >= 50 && numFmtId <= 58) {
            return true;
        }
        if (formatCode == null || formatCode.isBlank()) {
            return false;
        }
        String lc = formatCode.toLowerCase(Locale.ROOT);
        if (lc.contains("yy") || lc.contains("dd")) {
            return true;
        }
        return lc.contains("/") && lc.contains("m");
    }

    static boolean looksLikePercentFormat(String formatCode) {
        return formatCode != null && formatCode.contains("%");
    }

    static boolean looksLikeGroupedFormat(String formatCode) {
        if (formatCode == null) {
            return false;
        }
        return formatCode.contains("#,##") || formatCode.contains("#,#");
    }
}
