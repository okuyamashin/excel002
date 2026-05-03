package jp.engawa.excel002;

import static java.nio.charset.StandardCharsets.UTF_8;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

final class SheetTsvRowDeletionTest {

    private static final String SHEETS_ONE =
            "{\"file_name\":\"x\",\"last_modified\":0,\"file_md5\":\"m\",\"sheets_size\":1,\"sheets\":[{\"sheet_name\":\"S\",\"sheet_index\":0,\"max_row_num\":3,\"max_column_num\":2,\"sheet_id\":\"9\",\"relationship_id\":\"r\",\"worksheet_xml\":\"xl/w.xml\"}]}";

    @Test
    void remapMergeRefsAfterDeletingMiddleRow(@TempDir Path root) throws Exception {
        Path data = root.resolve("data");
        Files.createDirectories(data);
        Files.writeString(data.resolve("sheets.json"), SHEETS_ONE);
        writePlainGridsExceptMerge(data);
        Files.writeString(
                data.resolve("9.merge.tsv"),
                """
                0,0\t0,0
                0,0\t0,0
                (1,3)\t0,0
                """);

        SheetTsvRowDeletion.deleteRows(data, "9", 2);

        String jsonAfter = Files.readString(data.resolve("sheets.json"), UTF_8);
        Assertions.assertTrue(jsonAfter.contains("\"max_row_num\":2"), jsonAfter);

        var mergeLines = Files.readAllLines(data.resolve("9.merge.tsv"), UTF_8);
        String mergeLeft = mergeLines.get(mergeLines.size() - 1).split("\t", -1)[0].trim();
        Assertions.assertEquals("(1,2)", mergeLeft);
        Assertions.assertFalse(Files.isRegularFile(data.resolve("9.value.tsv.sheetrowdel.tmp")));
    }

    @Test
    void refuseWhenMergeAnchorWouldBeRemoved(@TempDir Path root) throws Exception {
        Path data = root.resolve("data");
        Files.createDirectories(data);
        Files.writeString(data.resolve("sheets.json"), SHEETS_ONE);
        writePlainGridsExceptMerge(data);
        Files.writeString(
                data.resolve("9.merge.tsv"),
                """
                0,0\t0,0
                0,1\t(1,2)
                0,0\t0,0
                """);

        IOException thrown = Assertions.assertThrows(IOException.class, () -> SheetTsvRowDeletion.deleteRows(data, "9", 2));
        Assertions.assertTrue(thrown.getMessage().contains("削除不能"));

        Assertions.assertFalse(Files.isRegularFile(data.resolve("9.value.tsv.sheetrowdel.tmp")));
        Assertions.assertTrue(Files.readString(data.resolve("sheets.json"), UTF_8).contains("\"max_row_num\":3"));
    }

    private static void writePlainGridsExceptMerge(Path data) throws Exception {
        String body = grid3x2("");
        Files.writeString(data.resolve("9.value.tsv"), body);
        Files.writeString(data.resolve("9.type.tsv"), body);
        Files.writeString(data.resolve("9.string.tsv"), body);
        Files.writeString(data.resolve("9.formula.tsv"), body);
        Files.writeString(data.resolve("9.format.tsv"), body);
        Files.writeString(data.resolve("9.style.tsv"), body);
        Files.writeString(data.resolve("9.formatted_value.tsv"), body);
    }

    private static String grid3x2(String cell) {
        String line = cell + "\t" + cell;
        return line + "\n" + line + "\n" + line + "\n";
    }
}
