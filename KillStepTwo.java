package cn.wangkf.xnan;

import cn.wangkf.util.ExcelUtils;
import com.google.common.collect.Lists;
import jxl.Sheet;
import jxl.Workbook;
import lombok.Data;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;


/**
 * Created by stanley.wang on 2019/12/25.
 */
public class KillStepTwo {

    private final static Double DEVIATION_VALUE = 0.015;

    private final static String FLODER_PATH = "E:\\data\\stepone\\";

    private final static String RESULT_FILE_NAME = "E:\\data\\stepone\\result.xls";

    private static List<String> f = Arrays.asList("CK", "CK15", "CK30", "HT15", "HT30", "LT15", "LT30");

    public static void main(String[] args) throws Exception {
        // step1
        List<StepOne> stepOnes = Lists.newArrayList();
        for (int i=0; i<f.size(); i++) {
            try {
                stepOnes.addAll(getSheetData(FLODER_PATH + f.get(i) + ".xls", i+1));
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        // step2
        dealSheetData(stepOnes);
        // step3
        transToExcel(stepOnes);
    }

    /**
     * 根据对象list生成excel
     * @param stepOnes
     */
    private static void transToExcel(List<StepOne> stepOnes) {
        List<List<String>> targetList = Lists.newArrayList();

        //添加表头
        targetList.add(Arrays.asList("时间坑", "数据源", "保留时间平均值", "物质名称", "L", "M", "N", "J"));

        //添加表数据
        stepOnes.forEach(s -> {
            targetList.add(s.toList());
        });

        //生成excel
        ExcelUtils.object2excel(targetList, RESULT_FILE_NAME, 0);
    }


    /**
     * 数据处理：排序过滤等
     * @param stepOnes
     */
    private static void dealSheetData(List<StepOne> stepOnes) {
        // 按照保留时间排序
        Collections.sort(stepOnes);

        // 按照误差值分坑
        int pitNum = 1; // 时间坑初始值
        double lastTime = Double.valueOf(stepOnes.get(0).reserveTime.replace(" ",""));
        for (int i=0; i< stepOnes.size(); i++) {
            double currentTime = Double.valueOf(stepOnes.get(i).reserveTime.replace(" ",""));
            if (Math.abs(currentTime-lastTime) > DEVIATION_VALUE) {
                pitNum++;
            }
            StepOne stepOne = stepOnes.get(i);
            stepOne.setPit(pitNum);
            stepOnes.set(i, stepOne);

            lastTime = currentTime;
        }
    }

    /**
     * 从Excel获取数据并组装数据
     * @param filePath
     * @param type
     * @return
     * @throws Exception
     */
    public static List<StepOne> getSheetData(String filePath, int type) throws Exception {
        List<StepOne> rowStrs = Lists.newArrayList();
        System.out.println(filePath);
        Workbook book = Workbook.getWorkbook(new File(filePath));
        //获得excel文件的sheet表
        Sheet sheet = book.getSheet(0);
        int rows = sheet.getRows();
        
        for (int i=1; i<rows; i++) {
            String timeInfo = new String(sheet.getCell(0, i).getContents());
            String jColInfo = new String(sheet.getCell(1, i).getContents());
            String nameInfo = new String(sheet.getCell(2, i).getContents());
            if (StringUtils.isNotBlank(timeInfo)) {
                StepOne stepOne = new StepOne();
                stepOne.setType(type);
                stepOne.setReserveTime(timeInfo);
                stepOne.setName(nameInfo.trim());
                stepOne.setLCol(sheet.getCell(3, i).getContents().replace(" ", ""));
                stepOne.setMCol(sheet.getCell(4, i).getContents().replace(" ", ""));
                stepOne.setNCol(sheet.getCell(5, i).getContents().replace(" ", ""));

                List<String> jCol = Lists.newArrayList();
                jCol.add(jColInfo.replace(" ", ""));
                stepOne.setJCol(jCol);
                rowStrs.add(stepOne);
            } else if (StringUtils.isNotBlank(jColInfo) ) {
                StepOne stepOne = rowStrs.get(rowStrs.size()-1);
                List<String> jCol = rowStrs.get(rowStrs.size()-1).getJCol();
                jCol.add(jColInfo.replace(" ", ""));
                stepOne.setJCol(jCol);
            }
        }
        return  rowStrs;
    }
}

@Data
class StepOne implements Comparable<StepOne>{
    int pit;
    int type;
    String reserveTime;
    String name;
    String lCol;
    String mCol;
    String nCol;
    List<String> jCol = Lists.newArrayList();

    @Override
    public int compareTo(StepOne o) {
        Double lTime = Double.valueOf(this.getReserveTime().replace(" ", ""));
        Double rTime = Double.valueOf(o.getReserveTime().replace(" ", ""));
        return ExcelUtils.compareTo(lTime, rTime);
    }

    public List<String> toList() {
        List<String> list = Lists.newArrayList();
        list.add(String.valueOf(pit));
        list.add(String.valueOf(type));
        list.add(reserveTime);
        list.add(name);
        list.add(lCol);
        list.add(mCol);
        list.add(nCol);
        list.add(jCol.toString());
        return list;
    }

}