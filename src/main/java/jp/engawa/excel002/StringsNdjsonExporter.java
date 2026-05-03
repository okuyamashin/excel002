package jp.engawa.excel002;

import java.io.BufferedWriter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/**
 * {@code xl/sharedStrings.xml} を NDJSON（{@code data/strings.ndjson}）に書き出す。
 * 各行は {@code schema_version}, {@code id}（0 始まりインデックスの文字列）, {@code value}（プレーンテキスト）を含む。
 */
final class StringsNdjsonExporter {
    private StringsNdjsonExporter() {}

    /**
     * {@code stringsNdjsonFile} に 1 行 1 オブジェクトを書く（{@code si} の出現順＝{@code id} 昇順）。
     * {@code sharedStrings.xml} が無いときは空ファイル。
     */
    static void write(Path excelRoot, Path stringsNdjsonFile) throws IOException {
        Path xl = excelRoot.resolve("xl");
        Path sharedXml = xl.resolve("sharedStrings.xml");
        Path parent = stringsNdjsonFile.toAbsolutePath().normalize().getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        List<String> texts = SheetValueTsvExporter.readSharedStrings(sharedXml);

        try (BufferedWriter w = Files.newBufferedWriter(stringsNdjsonFile, StandardCharsets.UTF_8)) {
            for (int i = 0; i < texts.size(); i++) {
                if (i > 0) {
                    w.write('\n');
                }
                String id = Integer.toString(i);
                w.write("{\"schema_version\":1,\"id\":");
                writeJsonString(w, id);
                w.write(",\"value\":");
                writeJsonString(w, texts.get(i));
                w.write('}');
            }
            if (!texts.isEmpty()) {
                w.write('\n');
            }
        }
    }

    private static void writeJsonString(Appendable w, String s) throws IOException {
        if (s == null) {
            w.append("\"\"");
            return;
        }
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
