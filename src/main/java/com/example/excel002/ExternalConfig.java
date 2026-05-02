package com.example.excel002;

import java.io.IOException;
import java.io.Reader;
import java.net.URISyntaxException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Properties;

/**
 * fat JAR と同じディレクトリにある {@code excel002.properties} を読み込む。
 * IDE や {@code target/classes} から起動した場合はカレントディレクトリを基準にする。
 */
public final class ExternalConfig {

    public static final String FILENAME = "excel002.properties";

    private ExternalConfig() {}

    public static Path configDirectory(Class<?> anchor) throws URISyntaxException {
        var source = anchor.getProtectionDomain().getCodeSource();
        if (source == null) {
            return Paths.get("").toAbsolutePath();
        }
        Path codePath = Paths.get(source.getLocation().toURI());
        if (Files.isRegularFile(codePath) && codePath.getFileName().toString().toLowerCase().endsWith(".jar")) {
            Path parent = codePath.getParent();
            return parent != null ? parent : Paths.get("").toAbsolutePath();
        }
        return Paths.get("").toAbsolutePath();
    }

    public static Properties load(Class<?> anchor) throws IOException, URISyntaxException {
        Path dir = configDirectory(anchor);
        Path file = dir.resolve(FILENAME);
        Properties properties = new Properties();
        if (!Files.isRegularFile(file)) {
            System.err.printf("設定ファイルが見つかりません（無視して続行）: %s%n", file.toAbsolutePath());
            return properties;
        }
        try (Reader reader = Files.newBufferedReader(file, StandardCharsets.UTF_8)) {
            properties.load(reader);
        }
        return properties;
    }
}
