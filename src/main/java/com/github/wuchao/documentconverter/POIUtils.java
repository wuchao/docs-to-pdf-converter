package com.github.wuchao.documentconverter;

import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;

import java.util.List;

public abstract class POIUtils {

    /**
     * 删除页眉页脚的超链接
     * https://stackoverflow.com/questions/35088893/apache-poi-remove-cthyperlink-low-level-code
     *
     * @param docx
     */
    public static void removeHyperlinkOfHeaderAndFooter(XWPFDocument docx) {

        List<XWPFHeader> headers = docx.getHeaderList();
        if (CollectionUtils.isNotEmpty(headers)) {
            headers.forEach(header -> {
                List<XWPFParagraph> paragraphs = header.getParagraphs();
                if (CollectionUtils.isNotEmpty(paragraphs)) {
                    paragraphs.forEach(paragraph -> {
                        removeHyperlinkOfParagraph((paragraph));
                    });
                }
            });
        }

        List<XWPFFooter> footers = docx.getFooterList();
        if (CollectionUtils.isNotEmpty(footers)) {
            footers.forEach(footer -> {
                List<XWPFParagraph> paragraphs = footer.getParagraphs();
                if (CollectionUtils.isNotEmpty(paragraphs)) {
                    paragraphs.forEach(paragraph -> {
                        removeHyperlinkOfParagraph((paragraph));
                    });
                }
            });
        }

    }

    public static void removeHyperlinkOfParagraph(XWPFParagraph paragraph) {
        List<CTHyperlink> hyperlinks = paragraph.getCTP().getHyperlinkList();
        if (CollectionUtils.isNotEmpty(hyperlinks)) {
            hyperlinks.forEach(hyperlink -> {
                paragraph.getCTP().removeHyperlink(0);
            });
        }
    }

}
