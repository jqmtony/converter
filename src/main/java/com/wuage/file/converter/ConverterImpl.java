package com.wuage.file.converter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.ConnectException;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.io.IOUtils;

import com.artofsolving.jodconverter.DefaultDocumentFormatRegistry;
import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.DocumentFormat;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import com.sun.star.uno.RuntimeException;
import com.sun.xml.internal.messaging.saaj.util.ByteOutputStream;

public class ConverterImpl {

    private final String                        host;
    private final int                           port;
    private final DefaultDocumentFormatRegistry formatReg = new DefaultDocumentFormatRegistry();
    private final Map<String, DocumentFormat>   formatMap = new HashMap<String, DocumentFormat>() {

                                                              private static final long serialVersionUID = 1L;

                                                              {
                                                                  // excel
                                                                  put("xlsx",
                                                                      formatReg.getFormatByFileExtension("xls"));
                                                                  put("xlsm",
                                                                      formatReg.getFormatByFileExtension("xls"));

                                                                  // doc
                                                                  put("doc",
                                                                      formatReg.getFormatByFileExtension("docx"));
                                                              }
                                                          };

    public ConverterImpl(String tcpUrl){
        String[] urlArray = tcpUrl.split(":");
        if (urlArray.length < 2) {
            throw new RuntimeException("tcpUrl地址不正确");
        }
        this.host = urlArray[0];
        this.port = Integer.parseInt(urlArray[1]);
    }

    public ByteOutputStream converter2Pdf(InputStream inputStream, String extension, int size) {
        try {
            if (size <= 0 || inputStream == null || inputStream.available() <= 0) {
                throw new RuntimeException("请求数据不可以为空");
            }
        } catch (IOException e2) {
            throw new RuntimeException("系统异常");
        }
        if (!ConverterUtil.needConverter(extension)) {
            throw new RuntimeException("该文件无法转换");
        }
        extension = extension.toLowerCase();
        ByteOutputStream outputStream = null;
        OpenOfficeConnection connection = null;
        File inputFile = null;
        File outputFile = null;
        DocumentFormat pdfFormat = formatReg.getFormatByFileExtension("pdf");
        DocumentFormat fromFormat = formatReg.getFormatByFileExtension(extension);
        if (fromFormat == null) {
            fromFormat = formatMap.get(extension);
        }
        if (fromFormat == null) {
            throw new RuntimeException("该文件无法转换");
        }
        try {
            inputFile = File.createTempFile("document", "." + extension);
            OutputStream inputFileStream = null;
            try {
                inputFileStream = new FileOutputStream(inputFile);
                IOUtils.copy(inputStream, inputFileStream);
            } finally {
                IOUtils.closeQuietly(inputFileStream);
            }
            outputFile = File.createTempFile("document", ".pdf");
            outputStream = new ByteOutputStream(size);
            // 开启连接
            connection = getConnection();
            // TODO 计算时间 超过3s打印日志
            DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
            converter.convert(inputFile, fromFormat, outputFile, pdfFormat);
            InputStream outputFileStream = null;
            try {
                outputFileStream = new FileInputStream(outputFile);
                IOUtils.copy(outputFileStream, outputStream);
            } finally {
                IOUtils.closeQuietly(outputFileStream);
            }
            return outputStream;
        } catch (Exception e) {
            if (outputStream != null) {
                outputStream.close();
            }
            e.printStackTrace();
            throw new RuntimeException("系统异常");
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
            if (inputFile != null) {
                inputFile.delete();
            }
            if (outputFile != null) {
                outputFile.delete();
            }
        }
    }

    private OpenOfficeConnection getConnection() throws ConnectException {
        OpenOfficeConnection connection = new SocketOpenOfficeConnection(host, port);
        connection.connect();
        return connection;
    }

}
