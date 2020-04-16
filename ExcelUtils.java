package cn.wangkf.util;

import jxl.Workbook;
import jxl.write.*;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.util.List;

import static jxl.format.Colour.YELLOW;

/**
 * Created by stanley.wang on 2019/12/30.
 */
public class ExcelUtils {

    /**
     *
     * 生成单个sheet页的excel
     * @param objectList 表头 + data
     * @param targetPath 导出文件路径
     * @param colorIndex 标记颜色的字段下标 (-1时表示没有颜色标记)
     */
    public static void object2excel(List<List<String>> objectList, String targetPath, int colorIndex) {
        WritableWorkbook book = null;
        WritableCellFormat wcf = null;
        try {
            File file = new File(targetPath);
            if (file.exists()) {
                file.delete();
            }

            book = Workbook.createWorkbook(file);
            WritableSheet sheet = book.createSheet("sheet1", 1);
            WritableFont font = new WritableFont(WritableFont.createFont("宋体"),
                    10, WritableFont.NO_BOLD);
            WritableCellFormat yellowFormat = new WritableCellFormat(font);
            yellowFormat.setBackground(YELLOW);
            WritableCellFormat whiteFormat = new WritableCellFormat(font);
            //whiteFormat.setBackground(WHITE);

            boolean yellowFlag = false;
            String lastValue = colorIndex >= 0 ? objectList.get(0).get(colorIndex) : "";

            for (int i = 0; i < objectList.size(); i++) {
                String currentValue = colorIndex >= 0 ? objectList.get(i).get(colorIndex) : "";
                if (!StringUtils.equals(currentValue, lastValue) && colorIndex >= 0) {
                    yellowFlag = yellowFlag ? false : true;
                }
                if (yellowFlag) {
                    wcf = yellowFormat;
                } else {
                    wcf = whiteFormat;
                }
                lastValue = currentValue;

                for (int j = 0; j < objectList.get(i).size(); j++) {
                    Label label = new Label(j, i, objectList.get(i).get(j), wcf);
                    sheet.addCell(label);
                }
            }

            book.write();
            book.close();
        } catch (Exception e) {
            e.printStackTrace();;
        }
    }

    /**
     * 生成多个sheet页的excel
     * @param sheets // sheet页 -> 行数据 -> 列数据
     * @param targetPath
     * @param colorIndex
     */
    public static void object2excel(List<List<List<String>>> sheets, String targetPath, List<Integer> colorIndex) {
        WritableWorkbook book = null;
        WritableCellFormat wcf = null;
        try {
            File file = new File(targetPath);
            if (file.exists()) {
                file.delete();
            }
            book = Workbook.createWorkbook(file);

            for (int z=0; z<sheets.size(); z++) {
                List<List<String>> objectList = sheets.get(z);
                WritableSheet sheet = book.createSheet("sheet1", z+1);
                WritableFont font = new WritableFont(WritableFont.createFont("宋体"),
                        10, WritableFont.NO_BOLD);
                WritableCellFormat yellowFormat = new WritableCellFormat(font);
                yellowFormat.setBackground(YELLOW);
                WritableCellFormat whiteFormat = new WritableCellFormat(font);
                //whiteFormat.setBackground(WHITE);

                boolean yellowFlag = false;
                String lastValue = colorIndex.get(z) >= 0 ? objectList.get(0).get(colorIndex.get(z)) : "";

                for (int i = 0; i < objectList.size(); i++) {
                    String currentValue = colorIndex.get(z) >= 0 ? objectList.get(i).get(colorIndex.get(z)) : "";
                    if (!StringUtils.equals(currentValue, lastValue) && colorIndex.get(z)>=0) {
                        yellowFlag = yellowFlag ? false : true;
                    }
                    if (yellowFlag) {
                        wcf = yellowFormat;
                    } else {
                        wcf = whiteFormat;
                    }
                    lastValue = currentValue;

                    for (int j = 0; j < objectList.get(i).size(); j++) {
                        Label label = new Label(j, i, objectList.get(i).get(j), wcf);
                        sheet.addCell(label);
                    }
                }
            }

            book.write();
            book.close();
        } catch (Exception e) {
            e.printStackTrace();;
        }
    }



    public static int compareTo(Double lTime, Double rTime) {
        if (lTime > rTime) {
            return 1;
        } else if (lTime < rTime) {
            return -1;
        } else {
            return 0;
        }
    }

}
