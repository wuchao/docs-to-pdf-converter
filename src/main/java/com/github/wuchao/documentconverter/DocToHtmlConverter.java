package com.github.wuchao.documentconverter;

import java.io.InputStream;
import java.io.OutputStream;

public class DocToHtmlConverter extends Converter {

    public DocToHtmlConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    @Override
    public void convert() throws Exception {

        loading();

        WordToHtml doc = new WordToHtml();
        doc.readDoc(inStream);

        processing();

        doc.getHTMLFile(outStream);

        finished();

    }

}
