package bc.bms.common.workbook;

import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.*;

/**
 * 工作簿色彩
 */
public final class ColorPicker {

    private ColorPicker() {

    }

    public static XSSFColor getColor(String htmlColor) {
        htmlColor = htmlColor.replaceAll("#","");

        Integer red = Integer.valueOf(htmlColor.substring(0, 2), 16);
        Integer green = Integer.valueOf(htmlColor.substring(2, 4), 16);
        Integer blue = Integer.valueOf(htmlColor.substring(4, 6), 16);

        Color color = new Color(red, green, blue);

        XSSFColor xssfColor = new XSSFColor(color);

        return xssfColor;
    }

}
