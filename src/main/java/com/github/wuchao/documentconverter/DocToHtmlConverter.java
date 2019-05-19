package com.github.wuchao.documentconverter;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.Iterator;

public class DocToHtmlConverter extends Converter {

    private String resourceDir;

    public DocToHtmlConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    public DocToHtmlConverter(String resourceDir, InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
        this.resourceDir = resourceDir;
    }

    /**
     * 转换效果不理想
     *
     * @throws Exception
     */
    /*@Override
    public void convert() throws Exception {

        loading();

        processing();

        convert(inStream, outStream, System.getProperty("user.dir") +
                File.separator + "images" + File.separator);

        finished();

    }*/

    /**
     * https://www.huberylee.com/2018/03/28/Java%E4%B8%ADMS-Word%E8%BD%ACHTML-PDF%E5%AE%9E%E7%8E%B0/
     *
     * @param inputStream  DOC 文件输入流
     * @param outputStream HTML 文件输出流
     * @param imageSaveDir 图片存储路径
     */
    /*public void convert(InputStream inputStream, OutputStream outputStream, final String imageSaveDir) throws Exception {
        HWPFDocument document = new HWPFDocument(inputStream);
        AbstractWordConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance()
                .newDocumentBuilder().newDocument());

        // 将文档中的图片抽取放入到指定位置
        wordToHtmlConverter.setPicturesManager(new PicturesManager() {
            @Override
            public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float
                    widthInches, float heightInches) {

                try {
                    FileOutputStream fileOutputStream = new FileOutputStream(imageSaveDir + suggestedName);
                    fileOutputStream.write(content);
                    fileOutputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }

                return imageSaveDir + suggestedName;
            }
        });

        wordToHtmlConverter.processDocument(document);
        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        transformer.setOutputProperty(OutputKeys.ENCODING, StandardCharsets.UTF_8.toString());
        transformer.setOutputProperty(OutputKeys.INDENT, "true");
        transformer.setOutputProperty(OutputKeys.METHOD, "html");
        DOMSource domSource = new DOMSource(wordToHtmlConverter.getDocument());

        transformer.transform(domSource, new StreamResult(outputStream));
    }*/

    /**
     * https://blog.csdn.net/vipqiangqiang/article/details/8664288
     *
     * @param inputfile
     */
    /*public void convert(String inputfile) {

    }*/
    @Override
    public void convert() throws Exception {

        loading();

        processing();

        String html = convert(inStream);
        if (StringUtils.isNotBlank(html)) {
            outStream.write(html.getBytes());
        }

        finished();

    }

    /**
     * https://blog.csdn.net/j1231230/article/details/80712531
     *
     * @param inputStream
     * @return
     * @throws IOException
     * @throws TransformerException
     * @throws ParserConfigurationException
     */
    public String convert(InputStream inputStream) throws IOException, TransformerException, ParserConfigurationException {
        HWPFDocument wordDocument = new HWPFDocument(inputStream);

        WordToHtmlConverter wordToHtmlConverter =
                new WordToHtmlConverter(DocumentBuilderFactory
                        .newInstance().newDocumentBuilder().newDocument());
        // 保存图片，并返回图片的相对路径
        wordToHtmlConverter.setPicturesManager((content, pictureType, name, width, height) -> {
            try {
                FileOutputStream out = new FileOutputStream(resourceDir + name);
                out.write(content);
                out.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
            return "images/" + name;
        });

        wordToHtmlConverter.processDocument(wordDocument);
        Document htmlDocument = wordToHtmlConverter.getDocument();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(out);
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "HTML");
        serializer.transform(domSource, streamResult);
        out.close();
        return adjustMeta(new String(out.toByteArray()));
    }

    public static String adjustMeta(String content) {
        // 将不标准的 <meta> 转换为标准的 <meta></meta>
        org.jsoup.nodes.Document doc = Jsoup.parse(content);
        Elements meta = doc.getElementsByTag("meta");
        Iterator<Element> iterator = meta.iterator();
        while (iterator.hasNext()) {
            Element next = iterator.next();
            next.html(next.html() + " ");
        }
        content = doc.html();
        return content;
    }

}
