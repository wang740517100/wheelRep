package cn.wangkf.xnan;

import cn.wangkf.util.ExcelUtils;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.*;
import lombok.Data;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.text.DecimalFormat;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

import static jxl.format.Colour.RED;
import static jxl.format.Colour.WHITE;
import static jxl.format.Colour.YELLOW;

/**
 * Created by stanley.wang on 2020/3/9.
 */

@Data
class StepFour implements Comparable<StepFour>{
    private String name;
    private List<String> CK0;
    private List<String> CK15;
    private List<String> CK30;
    private List<String> HT15;
    private List<String> HT30;
    private List<String> LT15;
    private List<String> LT30;

    @Override
    public int compareTo(StepFour o) {

        List<Double> thisList = Arrays.asList(Double.parseDouble(this.getCK0().get(7)),
                Double.parseDouble(this.getCK15().get(7)),
                Double.parseDouble(this.getCK30().get(7)),
                Double.parseDouble(this.getHT15().get(7)),
                Double.parseDouble(this.getHT30().get(7)),
                Double.parseDouble(this.getLT15().get(7)),
                Double.parseDouble(this.getLT30().get(7)));
        Collections.sort(thisList);

        List<Double> oList = Arrays.asList(Double.parseDouble(o.getCK0().get(7)),
                Double.parseDouble(o.getCK15().get(7)),
                Double.parseDouble(o.getCK30().get(7)),
                Double.parseDouble(o.getHT15().get(7)),
                Double.parseDouble(o.getHT30().get(7)),
                Double.parseDouble(o.getLT15().get(7)),
                Double.parseDouble(o.getLT30().get(7)));
        Collections.sort(oList);

        return ExcelUtils.compareTo(thisList.get(6), oList.get(6));
    }

}

public class KillStepFour {

    public static void main(String[] args) throws Exception {
        List<StepFour> rowStrs = getSheetData("E:\\data\\stepfour\\nann.xls");

        Collections.sort(rowStrs);

        System.out.println(rowStrs.size());
        List<StepFour> res = rowStrs.subList(0, 11260);
//        List<StepFour> res = rowStrs.subList(11260, 21260);
//        List<StepFour> res = rowStrs.subList(21260, 35382);

        object2excel1(res, "E:\\data\\stepfour\\result.xls");

    }



    public static List<StepFour> getSheetData(String filePath) throws Exception {
        List<StepFour> rowStrs = Lists.newArrayList();
        Workbook book = Workbook.getWorkbook(new File(filePath));
        //获得excel文件的sheet表
        Sheet sheet = book.getSheet(0);
        int rows = sheet.getRows();

        Set<String> nameSet = Sets.newHashSet();
        for (int i=1; i<rows; i++) {
            String nameCell = getInfoByNum(sheet, i, 0);
            if (StringUtils.isNotBlank(nameCell)) {
                StepFour stepFour = new StepFour();
                stepFour.setName(nameCell);

                stepFour.setCK0(getList(Arrays.asList(getInfoByNum(sheet, i, 1),
                        getInfoByNum(sheet, i, 2), getInfoByNum(sheet, i, 3)), nameCell));
                stepFour.setCK15(getList(Arrays.asList(getInfoByNum(sheet, i, 4),
                        getInfoByNum(sheet, i, 5), getInfoByNum(sheet, i, 6)), nameCell));
                stepFour.setCK30(getList(Arrays.asList(getInfoByNum(sheet, i, 7),
                        getInfoByNum(sheet, i, 8), getInfoByNum(sheet, i, 9)), nameCell));
                stepFour.setHT15(getList(Arrays.asList(getInfoByNum(sheet, i, 10),
                        getInfoByNum(sheet, i, 11), getInfoByNum(sheet, i, 12)), nameCell));
                stepFour.setHT30(getList(Arrays.asList(getInfoByNum(sheet, i, 13),
                        getInfoByNum(sheet, i, 14), getInfoByNum(sheet, i, 15)), nameCell));
                stepFour.setLT15(getList(Arrays.asList(getInfoByNum(sheet, i, 16),
                        getInfoByNum(sheet, i, 17), getInfoByNum(sheet, i, 18)), nameCell));
                stepFour.setLT30(getList(Arrays.asList(getInfoByNum(sheet, i, 19),
                        getInfoByNum(sheet, i, 20), getInfoByNum(sheet, i, 21)), nameCell));

                rowStrs.add(stepFour);
            }
        }
        return  rowStrs;
    }

    private static List<String> getList(List<String> list, String nameCell) {
        List<String> result = Lists.newArrayList();
        result .add(list.get(0));
        result .add(list.get(1));
        result .add(list.get(2));
        Double a = Double.parseDouble(list.get(0).trim());
        Double b = Double.parseDouble(list.get(1).trim());
        Double c = Double.parseDouble(list.get(2).trim());


        Double avg = 0.0 , ste = 0.0, rs = 0.0;
        if (a + b + c > 0) {
            avg = (a + b + c)/3;
            ste = variance(a, b, c, avg);
            rs = new Double(ste/avg*100);
        }
        result.add(avg.toString());
        result.add(ste.toString());
        result.add(rs.toString());
        if (a + b + c == 0) {
            result.add("false1");
            result.add("4");
        } else {
            result.add("false");
            result.add("1");
        }


        // 如果rs > 20进入重计算
        if (rs > 20) {
            if (StringUtils.equals("100037760", nameCell)) {
//                System.out.println("--------------" + "nameCell：" + nameCell + ", " + a + ", " + b + ", " + c + "--------------------");
//                System.out.println(reCal(a, b, c));
            }
            return reCal(a, b, c);
        }



        return result;
    }

    private static List<String> reCal(Double a, Double b , Double c) {



        Double avg = (a + b + c)/3;

        Double avg1 = ((b+c)/2 + b + c)/3;
        Double avg2 = (a + (a+c)/2 + c)/3;
        Double avg3 = (a + b + (a+b)/2)/3;

        Double ste1 = variance((b+c)/2, b, c, avg1);
        Double rs1 = new Double(ste1/avg1*100);

        Double ste2 = variance(a, (a+c)/2, c, avg2);
        Double rs2 = new Double(ste2/avg2*100);

        Double ste3 = variance(a, b, (a+b)/2, avg3);
        Double rs3 = new Double(ste3/avg3*100);
//        System.out.println("ste1：" + ste1 + ", rs1：" + rs1 + "，ste2：" + ste2 + ", rs2：" + rs2 + "，ste3：" + ste3 + ", rs3：" + rs3);




        List<String> result = Lists.newArrayList();
//        result.addAll(Arrays.asList(a.toString(), b.toString(), c.toString()));

        if (a+b==0 || a+c==0 || b+c==0) {
            result.addAll(Arrays.asList(a.toString(), b.toString(), c.toString()));
            result.addAll(Arrays.asList("0", "0", "0", "false1", "3"));
        } else {
            if (rs1 > rs2) {
                if (rs2 > rs3) {
                    result.addAll(Arrays.asList(a.toString(), b.toString(), avg3.toString()));
                    result.addAll(Arrays.asList(avg3.toString(), ste3.toString(), rs3.toString(), "true3"));
                } else {
                    result.addAll(Arrays.asList(a.toString(), avg2.toString(), c.toString()));
                    result.addAll(Arrays.asList(avg2.toString(), ste2.toString(), rs2.toString(), "true2"));
                }
            } else {
                if (rs1 > rs3) {
                    result.addAll(Arrays.asList(a.toString(), b.toString(), avg3.toString()));
                    result.addAll(Arrays.asList(avg3.toString(), ste3.toString(), rs3.toString(), "true3"));
                } else {
                    result.addAll(Arrays.asList(avg1.toString(), b.toString(), c.toString()));
                    result.addAll(Arrays.asList(avg1.toString(), ste1.toString(), rs1.toString(), "true1"));
                }
            }

            if (Double.parseDouble(result.get(5)) > 20) {
                result.add("5");
            } else {
                result.add("2");
            }
        }

//        System.out.println(result);
        return result;

    }



    public static Double variance(Double a, Double b, Double c, Double avg) {
        Double av = a - avg > 0 ? a-avg : avg-a;
        Double bv = b - avg > 0 ? b-avg : avg-b;
        Double cv = c - avg > 0 ? c-avg : avg-c;
        Double res = (Math.pow(av, 2) + Math.pow(bv, 2) + Math.pow(cv, 2))/(3-1);
        return Math.sqrt(res);
    }

    private static String getInfoByNum(Sheet sheet, int i, int j) {
        String result = new String(sheet.getCell(j, i).getContents());
        return result != null ? result.trim() : "";
    }




    public static void object2excel1(List<StepFour> rowStrs, String targetPath) {
        WritableWorkbook book = null;
        WritableCellFormat cf = null;
        try {
            File file = new File(targetPath);
            if (file.exists()) {
                file.delete();
            }

            book = Workbook.createWorkbook(file);
            WritableSheet sheet = book.createSheet("sheet1", 1);
            WritableFont font = new WritableFont(WritableFont.createFont("宋体"),
                    10, WritableFont.NO_BOLD);
            WritableCellFormat ycf = new WritableCellFormat(font);
            ycf.setBackground(YELLOW);
            WritableCellFormat wcf = new WritableCellFormat(font);
            wcf.setBackground(WHITE);
            WritableCellFormat rcf = new WritableCellFormat(font);
            rcf.setBackground(RED);

            int yIndex = 0;
            for (StepFour info : rowStrs) {

                if (yIndex % 2 == 0) {
                    cf = wcf;
                } else {
                    cf = ycf;
                }

                sheet.addCell(new Label(0, yIndex+1, info.getName(), cf));

                sheet.addCell(new Label(1, yIndex+1, info.getCK0().get(0), StringUtils.equals(info.getCK0().get(6), "true1") ? rcf:cf));
                sheet.addCell(new Label(2, yIndex+1, info.getCK0().get(1), StringUtils.equals(info.getCK0().get(6), "true2") ? rcf:cf));
                sheet.addCell(new Label(3, yIndex+1, info.getCK0().get(2), StringUtils.equals(info.getCK0().get(6), "true3") ? rcf:cf));

                sheet.addCell(new Label(4, yIndex+1, info.getCK15().get(0), StringUtils.equals(info.getCK15().get(6), "true1") ? rcf:cf));
                sheet.addCell(new Label(5, yIndex+1, info.getCK15().get(1), StringUtils.equals(info.getCK15().get(6), "true2") ? rcf:cf));
                sheet.addCell(new Label(6, yIndex+1, info.getCK15().get(2), StringUtils.equals(info.getCK15().get(6), "true3") ? rcf:cf));

                sheet.addCell(new Label(7, yIndex+1, info.getCK30().get(0), StringUtils.equals(info.getCK30().get(6), "true1") ? rcf:cf));
                sheet.addCell(new Label(8, yIndex+1, info.getCK30().get(1), StringUtils.equals(info.getCK30().get(6), "true2") ? rcf:cf));
                sheet.addCell(new Label(9, yIndex+1, info.getCK30().get(2), StringUtils.equals(info.getCK30().get(6), "true3") ? rcf:cf));

                sheet.addCell(new Label(10, yIndex+1, info.getHT15().get(0), StringUtils.equals(info.getHT15().get(6), "true1") ? rcf:cf));
                sheet.addCell(new Label(11, yIndex+1, info.getHT15().get(1), StringUtils.equals(info.getHT15().get(6), "true2") ? rcf:cf));
                sheet.addCell(new Label(12, yIndex+1, info.getHT15().get(2), StringUtils.equals(info.getHT15().get(6), "true3") ? rcf:cf));

                sheet.addCell(new Label(13, yIndex+1, info.getHT30().get(0), StringUtils.equals(info.getHT30().get(6), "true1") ? rcf:cf));
                sheet.addCell(new Label(14, yIndex+1, info.getHT30().get(1), StringUtils.equals(info.getHT30().get(6), "true2") ? rcf:cf));
                sheet.addCell(new Label(15, yIndex+1, info.getHT30().get(2), StringUtils.equals(info.getHT30().get(6), "true3") ? rcf:cf));

                sheet.addCell(new Label(16, yIndex+1, info.getLT15().get(0), StringUtils.equals(info.getLT15().get(6), "true1") ? rcf:cf));
                sheet.addCell(new Label(17, yIndex+1, info.getLT15().get(1), StringUtils.equals(info.getLT15().get(6), "true2") ? rcf:cf));
                sheet.addCell(new Label(18, yIndex+1, info.getLT15().get(2), StringUtils.equals(info.getLT15().get(6), "true3") ? rcf:cf));

                sheet.addCell(new Label(19, yIndex+1, info.getLT30().get(0), StringUtils.equals(info.getLT30().get(6), "true1") ? rcf:cf));
                sheet.addCell(new Label(20, yIndex+1, info.getLT30().get(1), StringUtils.equals(info.getLT30().get(6), "true2") ? rcf:cf));
                sheet.addCell(new Label(21, yIndex+1, info.getLT30().get(2), StringUtils.equals(info.getLT30().get(6), "true3") ? rcf:cf));


                yIndex ++;
            }

            book.write();
            book.close();
        } catch (Exception e) {
            e.printStackTrace();;
        }
    }




















    public static void object2excel(List<StepFour> rowStrs, String targetPath) {
        WritableWorkbook book = null;
        WritableCellFormat cf = null;
        try {
            File file = new File(targetPath);
            if (file.exists()) {
                file.delete();
            }

            book = Workbook.createWorkbook(file);
            WritableSheet sheet = book.createSheet("sheet1", 1);
            WritableFont font = new WritableFont(WritableFont.createFont("宋体"),
                    10, WritableFont.NO_BOLD);
            WritableCellFormat ycf = new WritableCellFormat(font);
            ycf.setBackground(YELLOW);
            WritableCellFormat wcf = new WritableCellFormat(font);
            wcf.setBackground(WHITE);
            WritableCellFormat rcf = new WritableCellFormat(font);
            rcf.setBackground(RED);

            int yIndex = 0;
            for (StepFour info : rowStrs) {
//                if (StringUtils.equals("100037760", info.getName())) {
//                    System.out.println(info);
//                }
                if (info.getCK0().size() != 8 || info.getCK15().size() != 8 || info.getCK30().size() != 8 || info.getHT15().size() != 8 ||
                        info.getHT30().size() != 8 || info.getLT15().size() != 8 || info.getLT30().size() != 8 ) {
                    System.out.println(info);
                }

                if (yIndex % 2 == 0) {
                    cf = wcf;
                } else {
                    cf = ycf;
                }

                sheet.addCell(new Label(0, yIndex+1, info.getName(), cf));

                setCellData(sheet, 1, yIndex, cf, 0, rcf, info.getCK0());
                setCellData(sheet, 6, yIndex, cf, 1, rcf, info.getCK15());
                setCellData(sheet, 11, yIndex, cf, 2, rcf, info.getCK30());
                setCellData(sheet, 16, yIndex, cf, 3, rcf, info.getHT15());
                setCellData(sheet, 21, yIndex, cf, 4, rcf, info.getHT30());
                setCellData(sheet, 26, yIndex, cf, 5, rcf, info.getLT15());
                setCellData(sheet, 31, yIndex, cf, 6, rcf, info.getLT30());


                yIndex += 3;
            }

            book.write();
            book.close();
        } catch (Exception e) {
            e.printStackTrace();;
        }
    }

    private static List<String> cellName = Arrays.asList("CK_0_1_FPKM",  	"CK_0_2_FPKM", 	"CK_0_3_FPKM",
            "CK_15_1_FPKM",  	"CK_15_2_FPKM", 	"CK_15_3_FPKM",
            "CK_30_1_FPKM",  	"CK_30_2_FPKM", 	"CK_30_3_FPKM",
            "HT_15_1_FPKM",  	"HT_15_2_FPKM", 	"HT_15_3_FPKM",
            "HT_30_1_FPKM",  	"HT_30_2_FPKM", 	"HT_30_3_FPKM",
            "LT_15_1_FPKM",  	"LT_15_2_FPKM", 	"LT_15_3_FPKM",
            "LT_30_1_FPKM",  	"LT_30_2_FPKM", 	"LT_30_3_FPKM");

    private static void setCellData(WritableSheet sheet, int xIndex, int yIndex, WritableCellFormat cf,
                                    int cellNameIndex, WritableCellFormat rcf, List<String> infoList) throws WriteException {
        sheet.addCell(new Label(xIndex, yIndex,      cellName.get(cellNameIndex * 3 + 0), cf));
        sheet.addCell(new Label(xIndex, yIndex+1, cellName.get(cellNameIndex * 3 + 1), cf));
        sheet.addCell(new Label(xIndex, yIndex+2, cellName.get(cellNameIndex * 3 + 2), cf));

        WritableCellFormat cf1=null, cf2=null, cf3=null,  cf4=null;
        if (StringUtils.equals(infoList.get(6), "true1")) {
            cf1 = rcf;
            cf2 = cf;
            cf3 = cf;
            cf4 = cf;
        } else if (StringUtils.equals(infoList.get(6), "true2")) {
            cf1 = cf;
            cf2 = rcf;
            cf3 = cf;
            cf4 = cf;
        } else if (StringUtils.equals(infoList.get(6), "true3")) {
            cf1 = cf;
            cf2 = cf;
            cf3 = rcf;
            cf4 = cf;
        }  else if (StringUtils.equals(infoList.get(6), "false1")) {
            cf1 = rcf;
            cf2 = rcf;
            cf3 = rcf;
            cf4 = rcf;
        } else {
            cf1 = cf;
            cf2 = cf;
            cf3 = cf;
            cf4 = cf;
        }

        Double rs = Double.parseDouble(infoList.get(5));
        if (rs > 20) {
            cf4 = rcf;
        }


        sheet.addCell(new Label(xIndex+1, yIndex,      infoList.get(0), cf1));
        sheet.addCell(new Label(xIndex+1, yIndex+1, infoList.get(1), cf2));
        sheet.addCell(new Label(xIndex+1, yIndex+2, infoList.get(2), cf3));

        sheet.addCell(new Label(xIndex+2, yIndex,   infoList.get(3), cf4));
        sheet.addCell(new Label(xIndex+3, yIndex, infoList.get(4), cf4));
        sheet.addCell(new Label(xIndex+4, yIndex, infoList.get(5), cf4));
    }






}
