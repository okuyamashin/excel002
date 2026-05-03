package jp.engawa.excel002.jdbc;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Properties;

/**
 * {@code excel002.properties} の {@code jdbc.*} を使って JDBC 接続を開く。
 */
public final class JdbcConnections {

    private JdbcConnections() {}

    /**
     * 必須: {@code jdbc.url}。任意: {@code jdbc.username}（既定 root）、{@code jdbc.password}（既定は空文字）。
     */
    public static Connection connect(Properties applicationProperties) throws SQLException {
        String url = applicationProperties.getProperty("jdbc.url");
        if (url == null || url.isBlank()) {
            throw new SQLException("jdbc.url が設定されていません");
        }
        String user = applicationProperties.getProperty("jdbc.username", "root");
        String password = applicationProperties.getProperty("jdbc.password", "");
        return DriverManager.getConnection(url, user, password);
    }
}
