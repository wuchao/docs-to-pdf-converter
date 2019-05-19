package com.github.wuchao.documentconverter;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import lombok.extern.slf4j.Slf4j;

import java.io.*;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;

@Slf4j
public class DocToPDFConverter extends Converter {

    private String officeHome;

    private File fromFile;

    private File toFile;

    private String fromFilePath;

    private String toFilePath;

    private String baseFolder;

    /**
     * // 格式大全：前缀对应以下方法的 toFormat 值
     * // 0:Microsoft Word 97 - 2003 文档 (.doc)
     * // 1:Microsoft Word 97 - 2003 模板 (.dot)
     * // 2:文本文档 (.txt)
     * // 3:文本文档 (.txt)
     * // 4:文本文档 (.txt)
     * // 5:文本文档 (.txt)
     * // 6:RTF 格式 (.rtf)
     * // 7:文本文档 (.txt)
     * // 8:HTML 文档 (.htm)(带文件夹)
     * // 9:MHTML 文档 (.mht)(单文件)
     * // 10:MHTML 文档 (.mht)(单文件)
     * // 11:XML 文档 (.xml)
     * // 12:Microsoft Word 文档 (.docx)
     * // 13:Microsoft Word 启用宏的文档 (.docm)
     * // 14:Microsoft Word 模板 (.dotx)
     * // 15:Microsoft Word 启用宏的模板 (.dotm)
     * // 16:Microsoft Word 文档 (.docx)
     * // 17:PDF 文件 (.pdf)
     * // 18:XPS 文档 (.xps)
     * // 19:XML 文档 (.xml)
     * // 20:XML 文档 (.xml)
     * // 21:XML 文档 (.xml)
     * // 22:XML 文档 (.xml)
     * // 23:OpenDocument 文本 (.odt)
     * // 24:WTF 文件 (.wtf)
     */
    private int toFormat;

    public DocToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    public DocToPDFConverter(String officeHome, File fromFile, File toFile, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(null, null, showMessages, closeStreamsWhenComplete);
        this.officeHome = officeHome;
        this.fromFile = fromFile;
        this.toFile = toFile;
    }

    public DocToPDFConverter(String fromFilePath, String toFilePath, int toFormat, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(null, null, showMessages, closeStreamsWhenComplete);
        this.fromFilePath = fromFilePath;
        this.toFilePath = toFilePath;
        this.toFormat = toFormat;
    }

    public DocToPDFConverter(String baseFolder, InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
        this.baseFolder = baseFolder;
    }

    /**
     * docx4j（转换效果不理想）
     *
     * @throws Exception
     */
//    @Override
//    public void convert() throws Exception {
//
//        loading();
//
//        InputStream iStream = inStream;
//
//        WordprocessingMLPackage wordMLPackage = Docx4jUtils.getMLPackage(iStream);
//
//        processing();
//
//        Docx4J.toPDF(wordMLPackage, outStream);
//
//        finished();
//
//    }


    /**
     * jodconverter（转换速度慢）
     * <p>
     * https://juejin.im/post/59eae4cd5188255e773bba5e
     * https://stackoverflow.com/questions/52460055/jodconverter-telling-me-office-manager-is-required-in-order-to-build-a-converter
     */
    /*@Override
    public void convert() {

        loading();

        processing();

        OfficeManager officeManager = LocalOfficeManager.builder()
                .install()
                .officeHome(officeHome)
                // 设置任务执行超时为5分钟，这里就是导致转换超时的配置
                .taskExecutionTimeout(1000 * 60 * 5L)
                // 设置任务队列超时为24小时
                .taskQueueTimeout(1000 * 60 * 60 * 24L)
                .build();

        try {
            // Start an office process and connect to the started instance (on port 2002).
            officeManager.start();
            // Convert
            JodConverter
                    .convert(fromFile)
                    .to(toFile)
                    .execute();

        } catch (OfficeException e) {
            e.printStackTrace();
        } finally {
            // Stop the office process
            OfficeUtils.stopQuietly(officeManager);
        }

        finished();

    }*/

    /**
     * jacob
     * 大文件比 jodconverter 转换效果好，67M 的文件，jodconverter 转换超时，jacob 转换时间为 40s
     * <p>
     * 需要把 jacob-1.19-x64.dll 放到 java.library.path 的任一路径中
     * System.getProperty("java.library.path") 的值:
     * D:\Program Files\Java\java-1.8.0-openjdk-1.8.0.181\bin;C:\WINDOWS\Sun\Java\bin;C:\WINDOWS\system32;C:\WINDOWS;C:\WINDOWS\system32;C:\WINDOWS;C:\WINDOWS\System32\Wbem;C:\WINDOWS\System32\WindowsPowerShell\v1.0\;C:\WINDOWS\System32\OpenSSH\;D:\Program Files\Git\cmd;%JAVA_HOME%\bin;%JAVA_HOME%\jre\bin;%GRADLE_HOME%\bin;D:\Program Files\Gradle\gradle-3.5\bin;%OPEN_JDK_HOME%\bin;%OPEN_JDK_HOME%\jre\bin;D:\Program Files\Gradle\gradle-4.9\bin;C:\Users\wu947\AppData\Local\Microsoft\WindowsApps;D:\Program Files\Microsoft\VS Code\bin;C:\Users\wu947\AppData\Local\Microsoft\WindowsApps;;.
     */
    /*@Override
    public void convert() {

        loading();

        processing();

        convertWordFmt(fromFilePath, toFilePath, toFormat);

        finished();
    }*/
    @Override
    public void convert() {

        loading();

        processing();

        finished();
    }

    /**
     * doc 和 docx 格式互相转换、doc/docx 转 pdf
     * 仅适用于 Windows 系统，并且需要安装 Windows Office / WPS
     * Windows Office 和 WPS（商用）都是收费的
     *
     * @param srcPath  源文件
     * @param descPath 目标文件
     * @param fmt      目标格式
     * @return
     */
    public static boolean convertWordFmt(String srcPath, String descPath, int fmt) {

        // 实例化 ComThread 线程与 ActiveXComponent
        ComThread.InitSTA();

        /**
         * 参数为 Word.Application：需要安装 Office
         * 参数为 kwps.application：需要安装 WPS
         */
        ActiveXComponent app = new ActiveXComponent("kwps.Application");

        try {
            // 文档隐藏时进行应用操作
            app.setProperty("Visible", new Variant(false));
            // 实例化模板 Document 对象
            Dispatch document = app.getProperty("Documents").toDispatch();
            // 打开 Document
            // 如何手动用 Office 打开文档，显示“受保护的视图，Office已检测到该文件存在问题。编辑此文件可能会损坏您的计算机”，
            // 那么这里打开也会报一下异常信息：
            // com.jacob.com.ComFailException: Invoke of: Open
            //
            // Source: Microsoft Office Word
            // Variant foo = Dispatch.call(document, "Open", srcPath, true, true);
            Variant foo = Dispatch.invoke(document, "Open", Dispatch.Method,
                    new Object[]{srcPath, new Variant(true), new Variant(true)}, new int[1]);
            if (foo.getvt() == 0) {
                return false;
            }
            Dispatch doc = foo.toDispatch();
            // 另存为
            // Dispatch.call(doc, "SaveAs", descPath, fmt);
            Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[]{descPath, new Variant(fmt)}, new int[1]);
            Dispatch.call(doc, "Close", new Variant(false));

            return true;

        } finally {
            // 释放线程与 ActiveXComponent
            app.invoke("Quit", new Variant[]{});
            ComThread.Release();
        }
    }

    /**
     * 貌似需要安装 micro office
     * https://github.com/Nabarun/DocToPdfConverter_AWSJava/blob/master/src/main/java/DocToPdfConversion.java
     *
     * @param inputStream
     * @param outputStream
     * @throws IOException
     */
    public void convertDocToPdf(String baseFolder, InputStream inputStream, OutputStream outputStream) throws IOException {
        IConverter converter = LocalConverter.builder()
                .baseFolder(new File(baseFolder))
                .workerPool(20, 25, 2, TimeUnit.SECONDS)
                .processTimeout(5, TimeUnit.SECONDS)
                .build();

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();

        Future<Boolean> conversion = converter
                .convert(inputStream).as(DocumentType.DOC)
                .to(byteArrayOutputStream).as(DocumentType.PDF)
                // optional
                .prioritizeWith(1000)
                .schedule();

        try {
            conversion.get();
        } catch (InterruptedException e) {
            e.printStackTrace();
        } catch (ExecutionException e) {
            e.printStackTrace();
        }

        byteArrayOutputStream.writeTo(outputStream);
        byteArrayOutputStream.flush();
        byteArrayOutputStream.close();

    }


    protected static String FONT = "src/main/resources/malgun.ttf";

    /**
     * 转换效果不好
     * https://github.com/bjsystems/doc2pdf/blob/master/src/main/java/com/bj/doc2pdf/convert/DocToPDFConverter.java
     *
     * @param inputStream
     * @param outputStream
     */
    /*public void convertDocToPdf(InputStream inputStream, OutputStream outputStream) {
        Document document = new Document();

        try {
            POIFSFileSystem fs = new POIFSFileSystem(inputStream);
            HWPFDocument doc = new HWPFDocument(fs);
            WordExtractor we = new WordExtractor(doc);

            PdfWriter writer = PdfWriter.getInstance(document, outputStream);

            Range range = doc.getRange();
            document.open();
            writer.setPageEmpty(true);
            document.newPage();
            writer.setPageEmpty(true);
            BaseFont bf = BaseFont.createFont(FONT, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

            String[] paragraphs = we.getParagraphText();
            for (int i = 0; i < paragraphs.length; i++) {

                org.apache.poi.hwpf.usermodel.Paragraph pr = range.getParagraph(i);
                // CharacterRun run = pr.getCharacterRun(i);
                // run.setBold(true);
                // run.setCapitalized(true);
                // run.setItalic(true);
                paragraphs[i] = paragraphs[i].replaceAll("\\cM?\r?\n", "");
                System.out.println("Length:" + paragraphs[i].length());
                System.out.println("Paragraph" + i + ": " + paragraphs[i].toString());

                // add the paragraph to the document
//                document.add(new Paragraph(paragraphs[i]));
                document.add(new Paragraph(paragraphs[i], new Font(bf, 12)));
            }

            System.out.println("Document testing completed");

        } catch (Exception e) {
            System.out.println("Exception during test");
            e.printStackTrace();
        } finally {
            // close the document
            document.close();
        }

    }*/

}
