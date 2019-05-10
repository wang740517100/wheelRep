package cn.wangkf.bak;

import com.google.common.collect.Lists;
import jxl.CellView;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

@Data
@AllArgsConstructor
class TxtFile {
    String index;
    String time;
    String percentage;
    String names;
}

public class Txt2Excel {

    private final static String FLODER_PATH = "E:\\myproject\\txt2xls\\";

    private static String c1Name="序号", c2Name="时间", c3Name="百分比(%)", c4Name="名称";

    public static void main(String args[]) {
        File file = new File(FLODER_PATH);
        String [] fileNames = file.list();

        Arrays.asList(fileNames).stream().forEach(p -> {
            if (p.contains(".txt")) {
                String txtFilePath = FLODER_PATH + p;
                System.out.println(txtFilePath);

                StringBuilder result = readTxt(txtFilePath);
                //System.out.println(result);

                List<TxtFile> txtFileList = dealTxt(result);
                //System.out.println(txtFileList);

                String xlsFilePath = StringUtils.substringBeforeLast(txtFilePath, ".") + ".xls";
                TransToExcel(txtFileList, xlsFilePath);

            }
        });
    }




    /**
     * 整理txt信息
     * @param result
     * @return
     */
    private static List<TxtFile> dealTxt(StringBuilder result){
        List<TxtFile> txtFileList = Lists.newArrayList();
        TxtFile txtVo;
        String time = "", percentage = "", names = "";
        int index = 0, flag = 1;
        while (true) {
            String tempStr = result.substring(index, result.indexOf("#", index));
            if (flag%3 == 1) {
                time = tempStr;
            } else if(flag%3 == 2) {
                percentage = tempStr;
            } else {
                List<String> strs = Arrays.asList(tempStr.split("\\$")).stream()
                        .filter(p -> StringUtils.isNoneBlank(p)).collect(Collectors.toList());
                names = strs.get(0);
                //names = tempStr;
                txtVo = new TxtFile(String.valueOf(flag/3), time, percentage, names);
                txtFileList.add(txtVo);
            }
            flag ++;

            index = result.indexOf("#", index) + 1;
            if (result.indexOf("#", index) == -1) {
                if (flag%3 == 0){
                    tempStr = result.substring(index, result.length());
                    List<String> strs = Arrays.asList(tempStr.split("\\$")).stream()
                        .filter(p -> StringUtils.isNoneBlank(p)).collect(Collectors.toList());
                    names = strs.get(0);
                    //names = tempStr;
                    txtVo = new TxtFile(String.valueOf(flag/3), time, percentage, names);
                    txtFileList.add(txtVo);
                }
                break;
            }
        }
        return txtFileList;
    }

    /**
     * 读取txt数据流
     * @param filePath
     * @return
     */
    private static StringBuilder readTxt(String filePath){
        //最终结果
        StringBuilder resultStr = new StringBuilder();
        //临时name字符串信息
        StringBuilder namesStr = new StringBuilder();
        File file = new File(filePath);
        BufferedReader reader = null;
        try{
            InputStreamReader isr = new InputStreamReader(new FileInputStream(file), "gb2312");
            reader = new BufferedReader(isr);
            String tempStr ;
            int index=1;
            while ((tempStr = reader.readLine()) != null) {
                //分解过滤行数据
                String[] tmp = tempStr.trim().split(" ");
                List<String> tmpStrs = Arrays.asList(tmp).stream().
                        filter(p -> StringUtils.isNoneBlank(p)).collect(Collectors.toList());

                //如果是空行则添加name字符串信息
                if (tmpStrs.size() == 0) {
                    resultStr.append(namesStr);
                    resultStr.append("#");
                    namesStr.delete(0, namesStr.length());
                    continue;
                }

                //通过第一个字符串是否是正整数且等于叠加数判断是不是新的一段信息
                if (tmpStrs.get(0).matches("^[+]{0,1}(\\d+)$") &&  //正整数
                        Integer.parseInt(tmpStrs.get(0)) == index) {
                    resultStr.append(tmpStrs.get(1));
                    resultStr.append("#");
                    resultStr.append(tmpStrs.get(2));
                    resultStr.append("#");
                    index++;
                } else {
                    //通过首字母是否大写判断是否是新的一条信息（不太准确）
                    char c = tmpStrs.get(0).replaceAll("\\d", "").charAt(0);
                    if (Character.isUpperCase(c)) {
                        namesStr.append("   $");
                        tmpStrs.stream().forEach(p -> {
                            //超过两位的字符串且前两位不为数字、包含字母的字符串
                            if ((p.trim().length()>2 && !p.trim().substring(0,2).matches("^[0-9]{1,}$"))
                                    || containsLetter(p.trim())) {
                                namesStr.append(p.trim() + " ");
                            }
                        });
                    } else {
                        //不是大写则接上一行信息
                        tmpStrs.stream().forEach(p -> {
                            if ((p.trim().length()>2 && !p.trim().substring(0,2).matches("^[0-9]{1,}$"))
                                    || containsLetter(p.trim())) {
                                namesStr.append(p.trim() + " ");
                            }
                        });
                    }
                }
            }
            //添加最后一个#后name的信息
            resultStr.append(namesStr);
            reader.close();
        } catch(Exception e) {
            e.printStackTrace();
        } finally{
            if (reader != null) {
                try{
                    reader.close();
                }catch(IOException e){
                }
            }
        }
        return resultStr;
    }

    public static boolean containsLetter(String strArg) {
        boolean isLetter = false;
        for (int i = 0; i < strArg.length(); i++) {
            if (Character.isLetter(strArg.charAt(i))) {
                isLetter = true;
                break;
            }
        }
        return isLetter;
    }

    /**
     * 将List以xls保存
     * @param txtFileList
     */
    private static void TransToExcel(List<TxtFile> txtFileList, String xlsFilePath) {
        WritableWorkbook book  = null;
        try {
            File file = new File(xlsFilePath);
            if (file.exists()) {
                file.delete();
            }
            // 创建一个xls文件、添加格式
            book = Workbook.createWorkbook(file);
            WritableSheet sheet = book.createSheet("info", 0);
            CellView cellView = new CellView();
            cellView.setAutosize(true);
            sheet.setColumnView(1, cellView );
            sheet.setColumnView(2, cellView );
            sheet.setColumnView(3, cellView );
            sheet.setColumnView(4, cellView );
            WritableCellFormat cellFormat = new WritableCellFormat();
            cellFormat.setAlignment(Alignment.CENTRE);
            cellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);

            Label label1 = new Label(0, 0, c1Name);
            label1.setCellFormat(cellFormat);
            Label label2 = new Label(1, 0, c2Name);
            label2.setCellFormat(cellFormat);
            Label label3 = new Label(2, 0, c3Name);
            label3.setCellFormat(cellFormat);
            Label label4 = new Label(3, 0, c4Name);
            label4.setCellFormat(cellFormat);
            sheet.addCell(label1);
            sheet.addCell(label2);
            sheet.addCell(label3);
            sheet.addCell(label4);
            //绑定数据
            for (int i = 0; i < txtFileList.size(); i++) {
                TxtFile txtVo = txtFileList.get(i);
                if (txtVo != null) {
                    Label index = new Label(0, (i+1), txtVo.getIndex());
                    Label time = new Label(1, (i+1), txtVo.getTime());
                    Label percentage= new Label(2, (i+1), txtVo.getPercentage());
                    Label name = new Label(3, (i+1), txtVo.getNames());
                    sheet.addCell(index);
                    sheet.addCell(name);
                    sheet.addCell(percentage);
                    sheet.addCell(time);
                }
            }
            book.write();
            book.close();
        } catch (Exception e) {
            e.printStackTrace();;
        }
    }
}