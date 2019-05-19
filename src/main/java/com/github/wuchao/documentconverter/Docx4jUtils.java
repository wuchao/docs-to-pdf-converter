package com.github.wuchao.documentconverter;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.CollectionUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.ObjectFactory;

import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.stream.Collectors;

/**
 * docx4j基本操作: https://www.cnblogs.com/cxxjohnson/p/7833911.html
 */
@Slf4j
public abstract class Docx4jUtils {

    private final static ObjectFactory objectFactory = Context.getWmlObjectFactory();

    private static final String DOC = ".doc";
    private static final String DOCX = ".docx";

    /*public static WordprocessingMLPackage getMLPackage(InputStream iStream) throws Exception {
        PrintStream originalStdout = System.out;

        //Disable stdout temporarily as Doc convert produces alot of output
        System.setOut(new PrintStream(new OutputStream() {
            @Override
            public void write(int b) {
                //DO NOTHING
            }
        }));

        // 最新的 Docx4J 已经废弃掉该方法了
        WordprocessingMLPackage mlPackage = Doc.convert(iStream);

        System.out.println(originalStdout);
        return mlPackage;
    }*/

    /**
     * 去除文档页眉和页脚的链接
     *
     * @param mlPackage
     */
    public static void removeHyperLinkOfHeaderAndFooter(WordprocessingMLPackage mlPackage) {
        HeaderPart headerPart = getHeaderPart(mlPackage);
        if (headerPart != null && CollectionUtils.isNotEmpty(headerPart.getContent())) {
            headerPart.getContent().stream()
                    .map(content -> {
                        if (content.toString().contains("HYPERLINK \"http")) {
                            return null;
                        } else {
                            return content;
                        }
                    })
                    .collect(Collectors.toList());
        }

        FooterPart footerPart = getFooterPart(mlPackage);
        if (footerPart != null && CollectionUtils.isNotEmpty(footerPart.getContent())) {
            footerPart.getContent().stream()
                    .forEach(content -> {
                        if (content.toString().contains("HYPERLINK \"http")) {
                            footerPart.setContents(null);
                            footerPart.setJAXBContext(null);
                        }
                    });

        }

    }

    public static HeaderPart getHeaderPart(WordprocessingMLPackage document) {
        RelationshipsPart relPart = document.getMainDocumentPart().getRelationshipsPart();
        Relationship rel = relPart.getRelationshipByType(Namespaces.HEADER);
        return (HeaderPart) relPart.getPart(rel);
    }

    public static FooterPart getFooterPart(WordprocessingMLPackage document) {
        RelationshipsPart relPart = document.getMainDocumentPart().getRelationshipsPart();
        Relationship rel = relPart.getRelationshipByType(Namespaces.FOOTER);
        return (FooterPart) relPart.getPart(rel);
    }

}
