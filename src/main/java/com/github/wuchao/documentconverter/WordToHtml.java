package com.github.wuchao.documentconverter;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.jsoup.Jsoup;
import org.jsoup.helper.W3CDom;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.Base64;
import java.util.HashMap;
import java.util.List;

/**
 * https://gitee.com/null_134_9802/codes/af72z8c6ok0b5lspgny4937
 * <p>
 * 使用的jar包：
 * jsoup-1.10.2.jar
 * ooxml-schemas-1.1.jar
 * openxml4j-1.0-beta.jar
 * org.apache.poi.xwpf.converter.core-1.0.6.jar
 * org.apache.poi.xwpf.converter.xhtml-1.0.6.jar
 * poi-3.13-20150929.jar
 * poi-ooxml-3.13-20150929.jar
 * poi-scratchpad-3.13-20150929.jar
 * xmlbeans-2.6.0.jar
 */
public class WordToHtml {

    Document document = null;

    public WordToHtml(File file) throws Exception {
        readFile(file);
    }

    public WordToHtml() {
    }

    /**
     * 读取文件
     *
     * @param file
     */
    public void readFile(File file) throws Exception {
        if (!file.exists()) {
            throw new RuntimeException("文件不存在：" + file.getPath());
        }

        if (file.getName().endsWith(".doc")) {
            readDoc(file);
        } else if (file.getName().endsWith(".docx")) {
            readDocx(file);
        } else {
            throw new RuntimeException("不是Word文档，程序无法处理");
        }
    }

    /**
     * 读取 doc 文件
     *
     * @param doc
     */
    public void readDoc(File doc) throws Exception {
        InputStream input = new FileInputStream(doc);
        readDoc(input);
    }

    /**
     * 读取 doc 文件
     *
     * @param inputStream
     */
    public void readDoc(InputStream inputStream) throws Exception {
        HWPFDocument wordDocument = new HWPFDocument(inputStream);

        // 将 doc 转换成 XML 的转换器
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        wordToHtmlConverter.setPicturesManager(
                (byte[] content, PictureType pictureType,
                 String suggestedName, float widthInches, float heightInches) -> {
                    return suggestedName;
                });
        wordToHtmlConverter.processDocument(wordDocument);

        // 获取 xml 文档
        document = wordToHtmlConverter.getDocument();
        // 获取文档上的所有图片
        List<Picture> pics = wordDocument.getPicturesTable().getAllPictures();
        if (pics != null) {
            processDocImg(pics, document);
        }

        inputStream.close();
    }

    /**
     * 读取 docx 文件
     *
     * @param docx
     */
    public void readDocx(File docx) throws IOException {
        if (!docx.exists()) {
            throw new RuntimeException("文件不存在：" + docx.getPath());
        }

        FileInputStream fis = new FileInputStream(docx);
        readDocx(fis);
    }

    /**
     * 读取 docx 文件
     *
     * @param inputStream
     */
    public void readDocx(InputStream inputStream) throws IOException {
        XWPFDocument xwpfDocument = new XWPFDocument(inputStream);

        // 创建 html 转换器并设置缩进
        XHTMLOptions options = XHTMLOptions.create().indent(4);

        // 获取 html 数据
        ByteArrayOutputStream baOut = new ByteArrayOutputStream();

        //// TODO: 依赖版本升级，导致类找不到（java.lang.NoClassDefFoundError: org/apache/poi/POIXMLDocumentPart）
        XHTMLConverter.getInstance().convert(xwpfDocument, baOut, options);

        document = new W3CDom().fromJsoup(Jsoup.parse(baOut.toString()));

        // 获取文档上的图片数据
        List<XWPFPictureData> allPictures = xwpfDocument.getAllPictures();
        if (allPictures.size() > 0) {
            processDocxImg(allPictures, document);
        }

        inputStream.close();
        baOut.close();
    }

    /**
     * 处理 docx 文档的图片
     *
     * @param pics
     * @param document
     */
    private void processDocxImg(List<XWPFPictureData> pics, Document document) {

        // 存放图片的 map,格式：图片名：图片所在的 img 标签对象
        HashMap<String, Element> hm = new HashMap();
        NodeList imgs = document.getElementsByTagName("img");
        for (int i = 0; i < imgs.getLength(); i++) {
            Element item = (Element) imgs.item(i);
            hm.put(item.getAttribute("src"), item);
        }

        // 将每张图片转换成Base64码
        for (int i = 0; i < pics.size(); i++) {
            XWPFPictureData picData = pics.get(i);

            String picType = picData.getPackagePart().getPartName().getExtension();
            String imgKey = picData.getPackagePart().getPartName().getName().substring(1);

            // 转成 Base64
            String pic64 = Base64.getEncoder().encodeToString(picData.getData());
            if (hm.get(imgKey) != null) {
                hm.get(imgKey).setAttribute("src", "data:image/" + picType + ";base64," + pic64);
            }
        }
    }

    /**
     * 处理 Doc 的图片
     * 将图片转换成 64 位编码，插入的文档中
     *
     * @param pics
     * @param document
     * @throws Exception
     */
    private void processDocImg(List<Picture> pics, Document document) throws Exception {
        // 存放图片的 map，格式：图片名：图片所在的 img 标签对象
        HashMap<String, Element> hm = new HashMap();
        NodeList imgs = document.getElementsByTagName("img");
        for (int i = 0; i < imgs.getLength(); i++) {
            Element item = (Element) imgs.item(i);
            hm.put(item.getAttribute("src"), item);
        }

        // 将每张图片转换成 Base64 码
        ByteArrayOutputStream baops = new ByteArrayOutputStream();
        for (int i = 0; i < pics.size(); i++) {
            Picture pic = pics.get(i);

            // 清空流
            baops.reset();
            // 将图片写入二进制流中
            pic.writeImageContent(baops);
            // 转成Base64
            String pic64 = Base64.getEncoder().encodeToString(baops.toByteArray());

            // 获取文件名
            String picName = pic.suggestFullFileName();

            //// TODO: 会出现空指针异常???
            if (hm.get(picName) != null) {
                hm.get(picName).setAttribute("src", "data:image/" + getPicType(picName) + ";base64," + pic64);
            }
        }
    }

    /**
     * 获取图片的类型，即后缀名
     *
     * @param picName
     * @return
     */
    private String getPicType(String picName) {
        if (picName != null) {
            int i = picName.lastIndexOf(".");
            return picName.substring(i + 1);
        } else {
            return "";
        }
    }

    /**
     * 将 XML 的 Document 对象转换成 String
     *
     * @param xmlResult
     */
    private void toFormatedXML(StreamResult xmlResult) {
        try {
            // 获取 XML 转换器
            Transformer transFormer = TransformerFactory.newInstance().newTransformer();

            // 设置转换的输出属性。
            transFormer.setOutputProperty(OutputKeys.METHOD, "html");
            transFormer.setOutputProperty(OutputKeys.INDENT, "yes");
            transFormer.setOutputProperty(OutputKeys.ENCODING, "utf-8");

            // XSLT 要求名称空间支持
            DOMSource domSource = new DOMSource(document);
            transFormer.transform(domSource, xmlResult);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取转换好的 HTML 文件
     *
     * @return
     */
    public void getHTMLFile(File file) {
        toFormatedXML(new StreamResult(file));
    }

    /**
     * 获取转换好的 HTML 文件
     *
     * @return
     */
    public void getHTMLFile(OutputStream outputStream) {
        toFormatedXML(new StreamResult(outputStream));
    }

    /**
     * 获取转换好的 HTML 文档的内容
     *
     * @return
     */
    public String getHTMLText() {
        StringWriter sw = null;
        try {
            sw = new StringWriter();
            toFormatedXML(new StreamResult(sw));
            return sw.toString();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                sw.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return null;
    }


    /*public static void main(String[] args) {
        File file = new File("C:\\Users\\Administrator\\Desktop\\1.docx");
        File file1 = new File("C:\\Users\\Administrator\\Desktop\\1.html");
        File file2 = new File("C:\\Users\\Administrator\\Desktop\\2.doc");
        File file3 = new File("C:\\Users\\Administrator\\Desktop\\2.html");
        if (file.isFile()) {
            Word2Html docx = new Word2Html(file);
            docx.getHTMLFile(file1);
            String htmlText = docx.getHTMLText();
            System.out.println(htmlText);
        }
        if (file2.isFile()) {
            Word2Html doc = new Word2Html(file2);
            doc.getHTMLFile(file3);
            String htmlText = doc.getHTMLText();
            System.out.println(htmlText);
        }
        System.out.println("-----------");
    }*/
}
