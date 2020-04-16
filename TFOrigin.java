package cn.wangkf.xnan;

import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.*;
import lombok.Data;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import static jxl.format.Colour.RED;
import static jxl.format.Colour.WHITE;
import static jxl.format.Colour.YELLOW;

/**
 * Created by stanley.wang on 2020/2/14.
 */
@Data
class TFOriginInfo  {

    private String name;

    // 7组 每组3个 共21个
    private List<String> xColunms;

    private List<String> yColunms;

}

public class TFOrigin {

    public static void main(String[] args) throws Exception {
        List<TFOriginInfo> rowStrs = getSheetData("E:\\data\\tf\\nan.xls");

        List<TFOriginInfo> errorInfo = rowStrs.stream()
                .filter(r -> r.getName()== null || r.getXColunms() == null || r.getYColunms() == null).collect(Collectors.toList());
        System.out.println("error------------"  + errorInfo);
        System.out.println("--------------allsize=" + rowStrs.size() + "-------errorsize=" + errorInfo.size());

        System.out.println("请检查下面物质【第一条信息】" + rowStrs.stream().filter(r -> r.getYColunms().stream().anyMatch(x -> StringUtils.equals(x.trim(), "#")))
                .map(TFOriginInfo::getName).collect(Collectors.toSet()));
        System.out.println("请检查下面物质【第二条信息】" + rowStrs.stream().filter(r -> r.getXColunms().stream().anyMatch(x -> StringUtils.equals(x.trim(), "#")))
                .map(TFOriginInfo::getName).collect(Collectors.toSet()));
        List<TFOriginInfo> result = rowStrs.stream()
                .filter(r -> r.getName()!= null && r.getXColunms() != null && r.getYColunms() != null).collect(Collectors.toList());
        object2excel(result, "E:\\data\\tf\\result.xls");
    }


    public static List<TFOriginInfo> getSheetData(String filePath) throws Exception {
        List<TFOriginInfo> rowStrs = Lists.newArrayList();
        Workbook book = Workbook.getWorkbook(new File(filePath));
        //获得excel文件的sheet表
        Sheet sheet = book.getSheet(0);
        int rows = sheet.getRows();

        Set<String> nameSet = Sets.newHashSet();
        for (int i=1; i<rows; i++) {
            String cellInfo = getInfoByNum(sheet, i, 0);
            if (isStr2Num(cellInfo)) {
                List<String> group = Lists.newArrayList();
                int index = 1;
                while (index <= 27) {
                    String str = getInfoByNum(sheet, i, index);
                    if (StringUtils.isNotBlank(str)) {
                        if (isStr2Float(str)) {
                            group.add(str);
                        } else {
                            group.add("#");
                        }
                    }
                    index ++;
                }

                if (group.size() > 0 && group.size() != 21) {
                    System.out.println("name:" + cellInfo + "-------" + (nameSet.contains(cellInfo)? "第二个":"第一个") + "-------" + group);
                    for (int j = 0; j < 21-group.size(); j++) {
                        group.add("#");
                    }
                }
                if (group.size() == 0 || StringUtils.isBlank(cellInfo)) {
                    continue;
                }

                if (nameSet.contains(cellInfo)) {
                    for (TFOriginInfo info : rowStrs) {
                        if (StringUtils.equals(cellInfo, info.getName())) {
                            info.setXColunms(group);
                        }
                    }

                } else {
                    TFOriginInfo info = new TFOriginInfo();
                    info.setName(cellInfo);
                    info.setYColunms(group);
                    rowStrs.add(info);
                    nameSet.add(cellInfo);
                }
            }
        }
        return  rowStrs;
    }

    private static String getInfoByNum(Sheet sheet, int i, int j) {
        String result = new String(sheet.getCell(j, i).getContents());
        return result != null ? result.trim() : "";
    }

    public static boolean isStr2Num(String str) {
        Pattern pattern = Pattern.compile("^[0-9]*$");
        Matcher matcher = pattern.matcher(str);
        return matcher.matches();

    }

    public static boolean isStr2Float(String str) {
        try{
            Double.parseDouble(str);
            return true;
        }catch(NumberFormatException e){
            return false;
        }
    }

    public static void object2excel(List<TFOriginInfo> rowStrs, String targetPath) {
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
            whiteFormat.setBackground(WHITE);
            WritableCellFormat rcf = new WritableCellFormat(font);
            rcf.setBackground(RED);

            int y = 1;
            for (TFOriginInfo info : rowStrs) {
                if (y % 2 == 0) {
                    wcf = yellowFormat;
                } else {
                    wcf = whiteFormat;
                }
                sheet.addCell(new Label(0, y+1, info.getName(), wcf));

                // 1 -4
                int x = 1, index = 0;
                for (int i = 0; i < 7; i++) {
                    sheet.addCell(new Label(x, y, info.getYColunms().get(index), equalToY(info, index) ? rcf:wcf));
                    sheet.addCell(new Label(x, y+1, info.getYColunms().get(index+1), equalToY(info, index+1) ? rcf:wcf));
                    sheet.addCell(new Label(x, y+2, info.getYColunms().get(index+2), equalToY(info, index+2) ? rcf:wcf));

                    sheet.addCell(new Label(x+1, y, info.getXColunms().get(index), equalToX(info, index) ? rcf:wcf));
                    sheet.addCell(new Label(x+2, y, info.getXColunms().get(index+1), equalToX(info, index+1) ? rcf:wcf));
                    sheet.addCell(new Label(x+3, y, info.getXColunms().get(index+2), equalToX(info, index+2) ? rcf:wcf));
                    x += 4;
                    index += 3;
                }
                y += 3;
            }

            book.write();
            book.close();
        } catch (Exception e) {
            e.printStackTrace();;
        }
    }

    private static boolean equalToY(TFOriginInfo info, int index) {
        return StringUtils.equals(info.getYColunms().get(index), "#");
    }

    private static boolean equalToX(TFOriginInfo info, int index) {
        return StringUtils.equals(info.getXColunms().get(index), "#");
    }

}
