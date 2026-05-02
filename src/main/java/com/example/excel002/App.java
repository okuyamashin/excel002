package com.example.excel002;

import java.io.IOException;
import java.net.URISyntaxException;
import java.util.Properties;

public final class App {
    public static void main(String[] args) {
        try {
            Properties config = ExternalConfig.load(App.class);
            System.out.println(config.getProperty("message", "excel002"));
        } catch (IOException | URISyntaxException e) {
            System.err.println(e.getMessage());
            System.exit(1);
        }
    }
}
