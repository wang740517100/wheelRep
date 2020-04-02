package cn.wangkf;



import sun.misc.BASE64Decoder;
import sun.misc.BASE64Encoder;

import java.io.*;

/**
 * Created by stanley.wang on 2020/4/2.
 */
public class pdfStream {

    public static void main(String[] args) throws Exception {


//        //将PDF格式文件转成base64编码 
//        String base64String = getPDFBinary("E:\\Elasticsearch技术解析与实战.pdf");
//        System.out.println(base64String);


        String base64String = fileRead();

        //将base64的编码转成PDF格式文件,，保存到C:/Users/Administrator/Desktop/昊达临时/test2.pdf
        base64StringToPDF(base64String, "E:\\result.pdf");

    }

    static BASE64Encoder encoder = new sun.misc.BASE64Encoder();
    static BASE64Decoder decoder = new sun.misc.BASE64Decoder();

    /**
     * 获取 pdf 的 base 流
     * @param filePath
     * @return
     */
    public static String getPDFBinary(String filePath) {
        FileInputStream fin = null;
        BufferedInputStream bin = null;
        ByteArrayOutputStream baos = null;
        BufferedOutputStream bout = null;
        try {
            //建立读取文件的文件输出流
            fin = new FileInputStream(new File(filePath));
            //在文件输出流上安装节点流（更大效率读取）
            bin = new BufferedInputStream(fin);
            // 创建一个新的 byte 数组输出流，它具有指定大小的缓冲区容量
            baos = new ByteArrayOutputStream();
            //创建一个新的缓冲输出流，以将数据写入指定的底层输出流
            bout = new BufferedOutputStream(baos);
            byte[] buffer = new byte[1024];
            int len = bin.read(buffer);
            while (len != -1) {
                bout.write(buffer, 0, len);
                len = bin.read(buffer);
            }
            //刷新此输出流并强制写出所有缓冲的输出字节，必须这行代码，否则有可能有问题
            bout.flush();
            byte[] bytes = baos.toByteArray();
            //sun公司的API
            return encoder.encodeBuffer(bytes).trim();
            //apache公司的API
            //return Base64.encodeBase64String(bytes);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fin.close();
                bin.close();
                //关闭 ByteArrayOutputStream 无效。此类中的方法在关闭此流后仍可被调用，而不会产生任何 IOException
                //baos.close();
                bout.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return null;
    }

    /**
     * 将 pdf 的 base 流转换为文件
     * @param base64sString
     * @param filePath
     */
    public static void base64StringToPDF(String base64sString, String filePath) {
        BufferedInputStream bin = null;
        FileOutputStream fout = null;
        BufferedOutputStream bout = null;
        try {
            //将base64编码的字符串解码成字节数组
            byte[] bytes = decoder.decodeBuffer(base64sString);
            //apache公司的API
            //byte[] bytes = Base64.decodeBase64(base64sString);
            //创建一个将bytes作为其缓冲区的ByteArrayInputStream对象
            ByteArrayInputStream bais = new ByteArrayInputStream(bytes);
            //创建从底层输入流中读取数据的缓冲输入流对象
            bin = new BufferedInputStream(bais);
            //指定输出的文件
            File file = new File(filePath);
            //创建到指定文件的输出流
            fout = new FileOutputStream(file);
            //为文件输出流对接缓冲输出流对象
            bout = new BufferedOutputStream(fout);

            byte[] buffers = new byte[1024];
            int len = bin.read(buffers);
            while (len != -1) {
                bout.write(buffers, 0, len);
                len = bin.read(buffers);
            }
            //刷新此输出流并强制写出所有缓冲的输出字节，必须这行代码，否则有可能有问题
            bout.flush();

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                bin.close();
                fout.close();
                bout.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    public static String fileRead() throws Exception {
        File file = new File("E:\\pdf流.txt");//定义一个file对象，用来初始化FileReader
        FileReader reader = new FileReader(file);//定义一个fileReader对象，用来初始化BufferedReader
        BufferedReader bReader = new BufferedReader(reader);//new一个BufferedReader对象，将文件内容读取到缓存
        StringBuilder sb = new StringBuilder();//定义一个字符串缓存，将字符串存放缓存中
        String s = "";
        while ((s = bReader.readLine()) != null) {//逐行读取文件内容，不读取换行符和末尾的空格
            sb.append(s + "\n");//将读取的字符串添加换行符后累加存放在缓存中
        }
        bReader.close();
        String str = sb.toString();
        System.out.println(str);
        return str;
    }




}
