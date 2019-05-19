package com.github.wuchao.documentconverter;

import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class DocxToHtmlConverter extends Converter {

    private String resourceDir;

    public DocxToHtmlConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    public DocxToHtmlConverter(String resourceDir, InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
        this.resourceDir = resourceDir;
    }

    @Override
    public void convert() throws IOException {

        loading();

//        XWPFDocument document = new XWPFDocument(inStream);
//
//        XHTMLOptions options = XHTMLOptions.create();

        processing();

//        XHTMLConverter.getInstance().convert(document, outStream, options);
        convert(resourceDir, inStream, outStream);

        finished();
    }

    /**
     * https://www.codeproject.com/Questions/742161/Java-Convert-DOC-to-PDF-or-HTML
     * https://blog.csdn.net/j1231230/article/details/80712531
     *
     * @param resourceDir
     * @param inputStream
     * @param outputStream
     */
    public void convert(String resourceDir, InputStream inputStream, OutputStream outputStream) throws IOException {
        XWPFDocument document = new XWPFDocument(inputStream);

        // Prepare XHTML options (here we set the IURIResolver to load images from a "word/media" folder)
        // .URIResolver(new FileURIResolver(new File("word/media")));
        XHTMLOptions options = XHTMLOptions.create();

        // 存放图片的文件夹
        options.setExtractor(new FileImageExtractor(new File(resourceDir)));
        // html 中图片的路径
//        options.URIResolver(new BasicURIResolver("images"));

        XHTMLConverter.getInstance().convert(document, outputStream, options);
    }

}
