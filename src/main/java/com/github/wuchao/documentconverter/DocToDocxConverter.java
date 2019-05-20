package com.github.wuchao.documentconverter;

import java.io.InputStream;
import java.io.OutputStream;

public class DocToDocxConverter extends Converter {

    public DocToDocxConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    @Override
    public void convert() {

        loading();

        processing();

        //// TODO:

        finished();

    }

}
