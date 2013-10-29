package bc.bms.common.workbook;

import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 工作簿字体
 */
public final class Font extends XSSFFont {

    public static final String DEFAULT_HEADER_FONT_NAME = "黑体";

    public static final String DEFAULT_FONT_NAME = "宋体";

    private Font(String name, int size) {
        super();
        setFontName(name);
        setFontHeight(size);
    }

    public static Font defaultHeaderFont(XSSFWorkbook workbook) {
        Font font = new Font(DEFAULT_HEADER_FONT_NAME, DEFAULT_FONT_SIZE);

        registerFont(font, workbook);

        return font;
    }

    public static Font defaultFont(XSSFWorkbook workbook) {
        Font font = new Font(DEFAULT_FONT_NAME, DEFAULT_FONT_SIZE);

        registerFont(font, workbook);

        return font;
    }

    public static Font newFont(String name, XSSFWorkbook workbook) {
        Font font = new Font(name, DEFAULT_FONT_SIZE);

        registerFont(font, workbook);

        return font;
    }

    public static Font newFont(String name, short size, XSSFWorkbook workbook) {
        Font font = new Font(name, size);

        registerFont(font, workbook);

        return font;
    }

    public static Font newFont(String name, short size, short color, XSSFWorkbook workbook) {
        Font font = new Font(name, size);
        font.setColor(color);

        registerFont(font, workbook);

        return font;
    }

    public static Font newFont(String name, short size, boolean border, boolean italic, short color, XSSFWorkbook workbook) {
        Font font = new Font(name, size);
        font.setBold(border);
        font.setItalic(italic);
        font.setColor(color);

        registerFont(font, workbook);

        return font;
    }

    private static void registerFont(Font font, XSSFWorkbook workbook) {
        font.registerTo(workbook.getStylesSource());
    }

}
