package com.github.wuchao.documentconverter;

import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.*;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public class DocxToPDFConverter extends Converter {

    public DocxToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    /**
     * 转换速度快，效果好，但是转换成 pdf 后原 word 表格边框丢失
     *
     * @throws Exception
     */
    @Override
    public void convert() throws Exception {

        loading();

        XWPFDocument document = new XWPFDocument(inStream);

        PdfOptions options = PdfOptions.create();

        processing();

        /**
         * 转换前需要将文档页眉页脚的超链接去掉
         * 使用 POIUtil.removeHyperlinkOfHeaderAndFooter() 删除超链接无效
         */

        // 去掉页眉的超链接
        List<XWPFHeader> headers = document.getHeaderList();
        if (CollectionUtils.isNotEmpty(headers)) {
            headers.stream().forEach(header -> {
                List<XWPFParagraph> paragraphs = header.getParagraphs();
                if (CollectionUtils.isNotEmpty(paragraphs)) {
                    paragraphs.forEach(paragraph -> {
                        List<XWPFRun> runs = paragraph.getRuns();
                        if (CollectionUtils.isNotEmpty(runs)) {
                            runs.forEach(run -> {
                               String runStr = run.toString();
                                if (runStr.contains("http://") ||
                                        runStr.contains("https://") ||
                                        run.toString().contains("www.")) {

                                    System.out.println(run.toString());
                                    int size = run.getCTR().sizeOfBrArray();
                                    for (int i = 0; i < size; i++){
                                        run.setText("", i);
                                    }
                                    System.out.println(run.toString());
                                }
                            });
                        }
                    });
                }
            });
        }

        // 去掉页脚的超链接
        List<XWPFFooter> footerList = document.getFooterList();
        if (CollectionUtils.isNotEmpty(footerList)) {
            footerList.stream().forEach(footer -> {
                List<XWPFParagraph> paragraphs = footer.getParagraphs();
                if (CollectionUtils.isNotEmpty(paragraphs)) {
                    for (XWPFParagraph paragraph : paragraphs) {
                        String runStr;
                        List<XWPFRun> runs = paragraph.getRuns();
                        if (CollectionUtils.isNotEmpty(runs)) {
                            for (XWPFRun run : runs) {
                                runStr = run.toString();
                                if (runStr.contains("http://") ||
                                        runStr.contains("https://") ||
                                        run.toString().contains("www.")) {

                                    System.out.println(run.toString());
                                    int size = run.getCTR().sizeOfBrArray();
                                    for (int i = 0; i < size; i++){
                                        run.setText("", i);
                                    }
                                    System.out.println(run.toString());
                                }
                            }
                        }
                    }
                }
            });
        }

        PdfConverter.getInstance().convert(document, outStream, options);

        finished();

    }
}