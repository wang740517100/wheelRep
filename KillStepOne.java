package cn.wangkf.xnan;

import cn.wangkf.util.ExcelUtils;
import com.google.common.collect.Lists;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.*;
import lombok.Data;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

/**
 * Created by stanley.wang on 2019/12/23.
 */
public class KillStepOne {

    private final static Double DEVIATION_VALUE = 0.015;

    private final static String FLODER_PATH = "E:\\data\\";

    public static void main(String[] args) throws Exception {
        Arrays.asList("CK15", "CK30", "HT15", "HT30", "LT15").forEach(f -> generateExcel(f));
    }

    private static void generateExcel(String fileName) {
        List<List<Info>> sheets = Lists.newArrayList();
        try {
            sheets.add(getSheetData(FLODER_PATH + fileName + "-1_AnalysisReport.xls", 1));
            sheets.add(getSheetData(FLODER_PATH + fileName + "-2_AnalysisReport.xls", 2));
            sheets.add(getSheetData(FLODER_PATH + fileName + "-3_AnalysisReport.xls", 3));

            List<Info> totalSheet = dealSheetData(sheets);

            transToExcel(sheets, totalSheet, FLODER_PATH + fileName + ".xls");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     *
     * 从Excel获取数据并组装数据
     * @param filePath
     * @param type
     * @return
     * @throws Exception
     */
    public static List<Info> getSheetData(String filePath, int type) throws Exception {
        List<Info> rowStrs = Lists.newArrayList();
        Workbook book = Workbook.getWorkbook(new File(filePath));
        //获得excel文件的sheet表
        Sheet sheet = book.getSheet(0);
        int rows = sheet.getRows();

        int sheetFlag = 0; // 记录表格主数据开关：0、1是关闭，2是打开
        boolean nameFlag = false; //记录物质名称的开关
        int nameIndex = 0; // 物质名称游标
        for (int i=1; i<rows; i++) {
            String cellInfo = new String(sheet.getCell(0, i).getContents());
            if (StringUtils.equals(cellInfo, "用户质谱图")) {
                sheetFlag = 0;
            }

            if (sheetFlag == 2 && StringUtils.isNotBlank(cellInfo)) {
                Info info = new Info();
                info.setPeak(new String(sheet.getCell(0, i).getContents()));
                info.setStartTime(new String(sheet.getCell(2, i).getContents()));
                info.setReserveTime(new String(sheet.getCell(5, i).getContents()));
                info.setEndTime(new String(sheet.getCell(8, i).getContents()));
                info.setPeakHigh(new String(sheet.getCell(11, i).getContents()));
                info.setArea(new String(sheet.getCell(15, i).getContents()));
                info.setAreaPercent(new String(sheet.getCell(19, i).getContents()));
                info.setType(type);
                rowStrs.add(info);
            }

            // 匹配质谱图物质名称和表格
            if (nameFlag) {
                if (nameIndex >= rowStrs.size()) {
                    System.out.println("excel" + type + "质谱图数量大于表格行：" +  nameIndex + " > " + rowStrs.size());
                    throw new Exception("质谱图数量大于表格行");
                }
                // 设置当前Info的物质名称
                Info info = rowStrs.get(nameIndex);
                info.setName(cellInfo);
                rowStrs.set(nameIndex, info);
                nameFlag = false;
                nameIndex ++;
            }

            if (StringUtils.equals(cellInfo, "积分峰列表")) {
                sheetFlag ++;
            } else if (StringUtils.equals(cellInfo, "峰")) {
                sheetFlag++;
            } else if (StringUtils.equals(cellInfo, "质谱图结构")) {
                nameFlag = true;
            }
        }
        return  rowStrs;
    }

    /**
     * 汇总筛选数据
     * @param sheets
     * @return
     */
    private static List<Info> dealSheetData(List<List<Info>> sheets) {
        List<Info> totalSheet =  Lists.newArrayList();
        for (List<Info> infos : sheets) {
            for (Info info : infos) {
                totalSheet.add(info);
            }
        }

        // 按照保留时间排序
        Collections.sort(totalSheet);

        // 按照误差值分坑
        int pitNum = 1; // 时间坑初始值
        double lastTime = Double.valueOf(totalSheet.get(0).reserveTime.replace(" ",""));
        for (int i=0; i< totalSheet.size(); i++) {
            double currentTime = Double.valueOf(totalSheet.get(i).reserveTime.replace(" ",""));
            if (Math.abs(currentTime-lastTime) > DEVIATION_VALUE) {
                pitNum++;
            }
            Info info = totalSheet.get(i);
            info.setPit(pitNum);
            totalSheet.set(i, info);

            lastTime = currentTime;
        }

        return totalSheet;
    }

    /**
     * 生成 Excel
     * @param sheets
     * @param totalSheet
     * @param xlsFilePath
     * @throws IOException
     * @throws WriteException
     */
    private static void transToExcel(List<List<Info>> sheets, List<Info> totalSheet, String xlsFilePath) throws IOException, WriteException {

        // sheet -> 行数据 -> 列数据
        List<List<List<String>>> totalList = Lists.newArrayList();

        // 第1-3页
        sheets.forEach(sheet -> {
            List<List<String>> targetList = Lists.newArrayList();
            //添加表头
            targetList.add(Arrays.asList("时间坑", "数据源", "峰", "开始", "保留时间", "结束", "峰高", "面积", "面积百分比", "物质名称"));
            //添加表数据
            sheet.forEach(s -> {
                targetList.add(s.toList());
            });
            totalList.add(targetList);
        });

        // 第4页
        List<List<String>> sheetFour = Lists.newArrayList();
        //添加表头
        sheetFour.add(Arrays.asList("时间坑", "数据源", "峰", "开始", "保留时间", "结束", "峰高", "面积", "面积百分比", "物质名称"));
        //添加表数据
        totalSheet.forEach(s -> {
            sheetFour.add(s.toList());
        });
        totalList.add(sheetFour);

        //生成excel
        ExcelUtils.object2excel(totalList, xlsFilePath, Arrays.asList(-1, -1, -1, 0));
    }
}

@Data
class Info implements Comparable<Info>{
    int pit;
    int type;
    String peak;
    String startTime;
    String reserveTime;
    String endTime;
    String peakHigh;
    String area;
    String areaPercent;
    String name;

    @Override
    public int compareTo(Info o) {
        Double lTime = Double.valueOf(this.getReserveTime().replace(" ", ""));
        Double rTime = Double.valueOf(o.getReserveTime().replace(" ", ""));
        return ExcelUtils.compareTo(lTime, rTime);
    }

    public List<String> toList() {
        List<String> list = Lists.newArrayList();
        list.add(String.valueOf(pit));
        list.add(String.valueOf(type));
        list.add(peak);
        list.add(name);
        list.add(startTime);
        list.add(reserveTime);
        list.add(endTime);
        list.add(peakHigh);
        list.add(area);
        list.add(areaPercent);
        list.add(name);
        return list;
    }
}