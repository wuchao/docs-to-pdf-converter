package com.github.wuchao.documentconverter;

import org.junit.Test;

import java.io.*;

public class ConvertTests {

    String officeHome = "D:\\Program Files\\LibreOffice";
    String fileDir = System.getProperty("user.dir") + File.separator + "files" + File.separator;

    @Test
    public void testDoc2Pdf() throws Exception {
        InputStream inputStream = new FileInputStream(fileDir + "DocStructures.doc");
        OutputStream outputStream = new FileOutputStream(new File(fileDir + System.currentTimeMillis() + ".pdf"));
        boolean shouldShowMessages = true;
        Converter converter = new DocToPDFConverter(inputStream, outputStream, shouldShowMessages, true);
        converter.convert();
    }

    @Test
    public void testDoc2Pdf2() throws Exception {
//        String inputfilepath = fileDir + "DocStructures.doc";
        String inputfilepath = fileDir + "JMETER轻松入门与实战.doc";
        Converter converter = new DocToPDFConverter(officeHome,
                new File(inputfilepath),
                new File(fileDir + System.currentTimeMillis() + ".pdf"),
                true,
                true);
        converter.convert();
    }

    @Test
    public void testDoc2Pdf3() throws Exception {
        Converter converter = new DocToPDFConverter(fileDir + "Python-Django.doc",
                fileDir + System.currentTimeMillis() + ".pdf",
                17,
                true,
                true);
        converter.convert();
    }

    @Test
    public void testDoc2Pdf4() throws Exception {
        Converter converter = new DocToPDFConverter(new FileInputStream(fileDir + "Python-Django.doc"),
                new FileOutputStream(fileDir + System.currentTimeMillis() + ".pdf"),
                true,
                true);
        converter.convert();
    }

    @Test
    public void testDoc2Pdf5() throws Exception {
        Converter converter = new DocToPDFConverter(new File(fileDir + "Python-Django.doc"),
                new File(fileDir + System.currentTimeMillis() + "1111111111111111111111111111111111111out.html"),
                true,
                true);
        converter.convert();
    }

    @Test
    public void testDocx2Pdf() throws Exception {
//        String inputfilepath = fileDir+ "DocStructures.docx";
        String inputfilepath = fileDir + "Python-Django.docx";
        InputStream inputStream = new FileInputStream(inputfilepath);
        OutputStream outputStream = new FileOutputStream(new File(fileDir + System.currentTimeMillis() + ".pdf"));
        boolean shouldShowMessages = true;
        Converter converter = new DocxToPDFConverter(inputStream, outputStream, shouldShowMessages, true);
        converter.convert();
    }

    @Test
    public void testDoc2Docx() throws Exception {
        String inputfilepath = fileDir + "DocStructures.doc";
        InputStream inputStream = new FileInputStream(inputfilepath);
        OutputStream outputStream = new FileOutputStream(new File(fileDir + System.currentTimeMillis() + ".docx"));
        boolean shouldShowMessages = true;
        Converter converter = new DocToDocxConverter(inputStream, outputStream, shouldShowMessages, true);
        converter.convert();
    }

    @Test
    public void testDoc2Html() throws Exception {
//        String inputfilepath = fileDir + "Python-Django.doc";
        String inputfilepath = fileDir + "DocStructures.doc";
        Converter converter = new DocToHtmlConverter(new FileInputStream(new File(inputfilepath)),
                new FileOutputStream(new File(fileDir + System.currentTimeMillis() + ".html")),
                true, true);
        converter.convert();
    }

    @Test
    public void testDocx2Html() throws Exception {
        String inputfilepath = fileDir + "Python-Django.docx";
        InputStream inputStream = new FileInputStream(inputfilepath);
        OutputStream outputStream = new FileOutputStream(new File(fileDir + System.currentTimeMillis() + ".html"));
        Converter converter = new DocxToHtmlConverter(inputStream, outputStream, true, true);
        converter.convert();
    }

}
