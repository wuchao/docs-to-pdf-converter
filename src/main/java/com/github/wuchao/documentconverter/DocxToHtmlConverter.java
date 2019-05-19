package com.github.wuchao.documentconverter;

import java.io.InputStream;
import java.io.OutputStream;

public class DocxToHtmlConverter extends Converter {

    public DocxToHtmlConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    @Override
    public void convert() throws Exception {

        loading();

        WordToHtml docx = new WordToHtml();
        docx.readDocx(inStream);

        processing();

        docx.getHTMLFile(outStream);

        finished();
    }

}
