package com.wuage.file.converter;

import java.io.File;
import java.io.IOException;

import com.artofsolving.jodconverter.DefaultDocumentFormatRegistry;
import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.DocumentFormat;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

public class Main {


    public static long office2PDF(String sourceFile, String destFile) {
        // String OpenOffice_HOME = "D:/Program Files/OpenOffice.org 3";// 这里是OpenOffice的安装目录,C:\Program Files
        // (x86)\OpenOffice 4
        // String OpenOffice_HOME = "C:/Program Files (x86)/OpenOffice 4/";
        // 在我的项目中,为了便于拓展接口,没有直接写成这个样子,但是这样是尽对没题目的
        // 假如从文件中读取的URL地址最后一个字符不是 '\'，则添加'\'
        // if (OpenOffice_HOME.charAt(OpenOffice_HOME.length() - 1) != '/') {
        // OpenOffice_HOME += "/";
        // }
        Process pro = null;

        try {
            File inputFile = new File(sourceFile);
            if (!inputFile.exists()) {
                return -1;// 找不到源文件, 则返回-1
            }
            // 如果目标路径不存在, 则新建该路径
            File outputFile = new File(destFile);
            if (!outputFile.getParentFile().exists()) {
                outputFile.getParentFile().mkdirs();
            }
            // // 启动OpenOffice的服务
            // String command = OpenOffice_HOME
            // + "program\\soffice.exe -headless
            // -accept=\"socket,host=127.0.0.1,port=8100;urp;StarOffice.ServiceManager\" -nofirststartwizard";
            // pro = Runtime.getRuntime().exec(command);
            // connect to an OpenOffice.org instance running on port 8100
            DefaultDocumentFormatRegistry formatReg = new DefaultDocumentFormatRegistry();

            DocumentFormat pdfFormat = formatReg.getFormatByFileExtension("pdf");

            DocumentFormat fromFormat = formatReg.getFormatByFileExtension("xls");
            OpenOfficeConnection connection = new SocketOpenOfficeConnection("127.0.0.1", 8100);
            // OpenOfficeConnection connection = new SocketOpenOfficeConnection(8100);
            connection.connect();

            // convert
            long begin=System.currentTimeMillis();
            DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
            converter.convert(inputFile,fromFormat, outputFile,pdfFormat);
            System.out.println("转化成功");
            long end=System.currentTimeMillis();
            // close the connection
            connection.disconnect();
            // 封闭OpenOffice服务的进程
            // pro.destroy();
            return end-begin;
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // pro.destroy();
        }

        return 1;
    }

    public static void main(String[] args) {
        Main main = new Main();
        String sourceFile = "D:/测试数据/29186947397435984.xlsx";
        String destFile = "D:/测试数据/29186947397435984.pdf";
        long time = main.office2PDF(sourceFile, destFile);
        System.out.println(time);
    }

}
