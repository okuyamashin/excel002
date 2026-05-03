package jp.engawa.excel002;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Locale;

/**
 * OOXML の格納値とスタイルから「それっぽい」表示文字列を返す。失敗や未対応は {@code valueText} そのもの。
 */
final class SimpleFormattedValue {
    private static final DateTimeFormatter DATE =
            DateTimeFormatter.ofPattern("yyyy-MM-dd", Locale.ROOT);
    private static final DateTimeFormatter DATE_TIME =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss", Locale.ROOT);

    private SimpleFormattedValue() {}

    static String format(String valueText, String typeLabel, int styleIndex, OoxmlStylesResolver styles) {
        String raw = valueText;
        if (raw == null || raw.isEmpty()) {
            return raw;
        }
        String type = typeLabel;
        if ("string".equals(type) || "boolean".equals(type) || "error".equals(type)) {
            return raw;
        }
        try {
            double n = Double.parseDouble(raw.trim());
            int sid = styleIndex;
            int numFmtId = styles.numFmtIdForCellStyle(sid);
            String code = styles.formatCodeForCellStyle(sid);

            if ("date".equals(type)
                    || OoxmlStylesResolver.looksLikeDateFormat(numFmtId, code)) {
                return formatExcelSerial(n);
            }
            if (OoxmlStylesResolver.looksLikePercentFormat(code)) {
                return formatPercent(n);
            }
            if (OoxmlStylesResolver.looksLikeGroupedFormat(code)) {
                return formatGrouped(n);
            }
            return formatGeneralNumber(n);
        } catch (Exception e) {
            return raw;
        }
    }

    /** Excel 1900 日付系シリアル（Windows 既定）を簡易変換。 */
    private static String formatExcelSerial(double serial) {
        double whole = Math.floor(serial);
        double frac = serial - whole;
        long days = (long) whole;
        if (days >= 60) {
            days--;
        }
        LocalDate date = LocalDate.of(1899, 12, 30).plusDays(days);
        if (frac <= 1e-12) {
            return date.format(DATE);
        }
        long nanos = Math.round(frac * 86_400_000_000_000L);
        LocalDateTime ldt = LocalDateTime.of(date, LocalTime.MIDNIGHT).plusNanos(nanos);
        return ldt.format(DATE_TIME);
    }

    private static String formatPercent(double n) {
        BigDecimal v = BigDecimal.valueOf(n).multiply(BigDecimal.valueOf(100));
        v = v.stripTrailingZeros();
        if (v.scale() < 0) {
            v = v.setScale(0, RoundingMode.UNNECESSARY);
        }
        return v.toPlainString() + "%";
    }

    private static String formatGrouped(double n) {
        BigDecimal v = BigDecimal.valueOf(n).stripTrailingZeros();
        String plain = v.toPlainString();
        int dot = plain.indexOf('.');
        String intPart = dot >= 0 ? plain.substring(0, dot) : plain;
        String frac = dot >= 0 ? plain.substring(dot) : "";
        boolean neg = intPart.startsWith("-");
        String digits = neg ? intPart.substring(1) : intPart;
        StringBuilder sb = new StringBuilder();
        if (neg) {
            sb.append('-');
        }
        int len = digits.length();
        for (int i = 0; i < len; i++) {
            if (i > 0 && (len - i) % 3 == 0) {
                sb.append(',');
            }
            sb.append(digits.charAt(i));
        }
        sb.append(frac);
        return sb.toString();
    }

    private static String formatGeneralNumber(double n) {
        BigDecimal v = BigDecimal.valueOf(n).stripTrailingZeros();
        return v.toPlainString();
    }
}
