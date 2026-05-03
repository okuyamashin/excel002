package jp.engawa.excel002;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

/**
 * {@code xl/styles.xml} を読み、{@code data/styles.ndjson} に {@code cellXfs} インデックス（{@code .style.tsv} の ID）ごとの
 * 復元用 {@code xf} XML 断片と {@code numFmtId} / {@code formatCode} を NDJSON（JSON Lines）で書き出す。
 */
final class StylesJsonExporter {
    private static final String NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private StylesJsonExporter() {}

    /**
     * {@code stylesNdjsonFile} に 1 行 1 オブジェクトの NDJSON を書く（{@code cellXfs} の並び＝{@code id} 昇順）。
     * 各行は {@code schema_version}, {@code id}, {@code numFmtId}, {@code formatCode}, {@code restore_xml} を含む。
     * {@code xl/styles.xml} が無い、または {@code cellXfs} が空のときは **空ファイル**（0 バイト）とする。
     */
    static void write(Path excelRoot, Path stylesNdjsonFile) throws IOException {
        Path xl = excelRoot.resolve("xl");
        Path stylesXml = xl.resolve("styles.xml");
        Path parent = stylesNdjsonFile.toAbsolutePath().normalize().getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        if (!Files.isRegularFile(stylesXml)) {
            Files.writeString(stylesNdjsonFile, "", StandardCharsets.UTF_8);
            return;
        }

        OoxmlStylesResolver resolver = OoxmlStylesResolver.load(stylesXml);
        List<XfRecord> xfs = readCellXfFragments(stylesXml);

        try (BufferedWriter w = Files.newBufferedWriter(stylesNdjsonFile, StandardCharsets.UTF_8)) {
            for (int i = 0; i < xfs.size(); i++) {
                String id = Integer.toString(i);
                XfRecord xf = xfs.get(i);
                String formatCode = resolver.formatCodeForCellStyle(i);
                if (i > 0) {
                    w.write('\n');
                }
                w.write("{\"schema_version\":1,\"id\":");
                writeJsonString(w, id);
                w.write(",\"numFmtId\":");
                w.write(Integer.toString(xf.numFmtId()));
                w.write(",\"formatCode\":");
                writeJsonString(w, formatCode);
                w.write(",\"restore_xml\":");
                writeJsonString(w, xf.restoreXml());
                w.write('}');
            }
            if (!xfs.isEmpty()) {
                w.write('\n');
            }
        }
    }

    private record XfRecord(int numFmtId, String restoreXml) {}

    private static List<XfRecord> readCellXfFragments(Path stylesXml) throws IOException {
        List<XfRecord> list = new ArrayList<>();
        XMLInputFactory f = XMLInputFactory.newInstance();
        try (InputStream in = Files.newInputStream(stylesXml)) {
            XMLStreamReader r = f.createXMLStreamReader(in);
            boolean inCellXfs = false;
            while (r.hasNext()) {
                int ev = r.next();
                if (ev == XMLStreamConstants.START_ELEMENT && NS.equals(r.getNamespaceURI())) {
                    String ln = r.getLocalName();
                    if ("cellXfs".equals(ln)) {
                        inCellXfs = true;
                    } else if ("xf".equals(ln) && inCellXfs) {
                        int numFmtId = parseNumFmtIdAttr(r);
                        String fragment = captureXfSubtreeXml(r);
                        list.add(new XfRecord(numFmtId, fragment));
                    }
                } else if (ev == XMLStreamConstants.END_ELEMENT
                        && NS.equals(r.getNamespaceURI())
                        && "cellXfs".equals(r.getLocalName())) {
                    inCellXfs = false;
                }
            }
            r.close();
        } catch (XMLStreamException e) {
            throw new IOException("styles.xml の cellXfs 解析に失敗しました", e);
        }
        return list;
    }

    private static int parseNumFmtIdAttr(XMLStreamReader r) {
        String n = r.getAttributeValue(null, "numFmtId");
        if (n == null || n.isBlank()) {
            return 0;
        }
        try {
            return Integer.parseInt(n.trim());
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    /**
     * 現在位置が {@code <xf>} の {@code START_ELEMENT} であるとき、対応する終了まで読み、単体で意味を持つ断片 XML を返す（ルート {@code xf}
     * に {@code xmlns} を付与）。
     */
    private static String captureXfSubtreeXml(XMLStreamReader r) throws XMLStreamException {
        StringBuilder sb = new StringBuilder();
        sb.append("<xf xmlns=\"").append(escapeXmlAttr(NS)).append("\"");
        for (int i = 0; i < r.getAttributeCount(); i++) {
            String ln = r.getAttributeLocalName(i);
            String vu = r.getAttributeValue(i);
            sb.append(' ').append(ln).append("=\"").append(escapeXmlAttr(vu)).append("\"");
        }
        sb.append('>');
        int depth = 1;
        while (depth > 0 && r.hasNext()) {
            int ev = r.next();
            if (ev == XMLStreamConstants.START_ELEMENT) {
                sb.append('<').append(r.getLocalName());
                for (int i = 0; i < r.getAttributeCount(); i++) {
                    String ln = r.getAttributeLocalName(i);
                    String vu = r.getAttributeValue(i);
                    sb.append(' ').append(ln).append("=\"").append(escapeXmlAttr(vu)).append("\"");
                }
                sb.append('>');
                depth++;
            } else if (ev == XMLStreamConstants.END_ELEMENT) {
                sb.append("</").append(r.getLocalName()).append('>');
                depth--;
            } else if (ev == XMLStreamConstants.CHARACTERS || ev == XMLStreamConstants.CDATA) {
                sb.append(escapeXmlText(r.getText()));
            }
        }
        return sb.toString();
    }

    private static String escapeXmlAttr(String s) {
        if (s == null || s.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder(s.length() + 8);
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            switch (c) {
                case '&':
                    sb.append("&amp;");
                    break;
                case '"':
                    sb.append("&quot;");
                    break;
                case '<':
                    sb.append("&lt;");
                    break;
                default:
                    sb.append(c);
                    break;
            }
        }
        return sb.toString();
    }

    private static String escapeXmlText(String s) {
        if (s == null || s.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder(s.length() + 8);
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            switch (c) {
                case '&':
                    sb.append("&amp;");
                    break;
                case '<':
                    sb.append("&lt;");
                    break;
                case '>':
                    sb.append("&gt;");
                    break;
                default:
                    sb.append(c);
                    break;
            }
        }
        return sb.toString();
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
