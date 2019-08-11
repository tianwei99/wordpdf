package com.test.word.wordpdf.controller;

import com.test.word.wordpdf.util.ShellUtil;
import com.test.word.wordpdf.util.WordUtil;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.jodconverter.DocumentConverter;
import org.jodconverter.document.DocumentFormat;
import org.jodconverter.job.ConversionJobWithRequiredSourceFormatUnspecified;
import org.jodconverter.office.OfficeException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;

@RestController
public class WordController
{
//    @Autowired
//    private IWordService wordService;

    @Autowired
    private DocumentConverter documentConverter;

    @RequestMapping("/libreOffice/wordToPdf")
    public void toPdf() throws OfficeException {
        File wfile = new File("/usr/local/output.docx");
        File pfile = new File("/usr/local/output.pdf");
        documentConverter.convert(wfile).to(pfile).execute();

        byte[] bytes = new byte[0];
        InputStream is = new ByteArrayInputStream(bytes);
        OutputStream os = new ByteArrayOutputStream();
        ConversionJobWithRequiredSourceFormatUnspecified convert = documentConverter.convert(is);

        DocumentFormat pdf = null;

//        new BasicDocumentFormatRegistry();
        convert.as(pdf).to(os);

        byte[] pdfBytes = new byte[0];
        try {
            os.write(pdfBytes);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @RequestMapping("/openPDF")
    public void openPDF(HttpServletResponse response) {
        try {
            String targetPath = System.getProperty("java.io.tmpdir") + File.separator;
//            String docxPath = "http://192.168.200.201:9093/M00/00/01/wKjIyVqQ34yAUO-uAAAtY6q46LU28.docx";
            String docxPath = "/usr/local/output.docx";
//            WordUtil.libreOfficeWordToPdf(docxPath, targetPath);
            File file = new File(docxPath);
            try {
                String osName = System.getProperty("os.name");
                String command;
                if (osName.contains("Windows")) {
                    command = "soffice --convert-to pdf  -outdir " + targetPath + " " + docxPath;
                } else {
                    command = "doc2pdf --output=" + targetPath + File.separator
                            + file.getName().replaceAll(".(?i)docx", ".pdf") + " " + docxPath;
                }
                String result = ShellUtil.executeCommand(command);
                if (!result.equals("") && !result.contains("writer_pdf_Export")) {
                    return;
                }
            } catch (Exception e) {
                e.printStackTrace();
                throw e;
            }

            String name = docxPath.substring(docxPath.lastIndexOf("/") + 1);
            String[] nameArr = name.split("\\.");

//            WordUtil.findPdf(response,targetPath + nameArr[0] + ".pdf");
            targetPath = targetPath + nameArr[0] + ".pdf";
            response.setContentType("application/pdf");
            response.setCharacterEncoding("utf-8");
            FileInputStream in = new FileInputStream(new File(targetPath));
            OutputStream out = response.getOutputStream();
            byte[] b = new byte[512];
            while ((in.read(b)) != -1) {
                out.write(b);
            }
            out.flush();
            in.close();
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @GetMapping(value = "/openOffice/wordToPdf")
    public void wordToPdf2()
    {
        String fileUrl = "C:\\Users\\xyyto\\Desktop\\output.doc";
        String outpath = "C:\\Users\\xyyto\\Desktop\\output2.pdf";

        WordUtil.officeToPDF(fileUrl, outpath);
    }

    @GetMapping(value = "/poi/xwpf/wordToPdf")
    public void wordToPdf()
    {
        URL url = this.getClass().getResource("/");
        String classesUrl = url.getPath();
//        String fileUrl = classesUrl + "template/test.docx";
        String fileUrl = "C:\\Users\\xyyto\\Desktop\\output.docx";

        String outpath = "C:\\Users\\xyyto\\Desktop\\output2.pdf";

        InputStream source;
        OutputStream target;
        try {
            source = new FileInputStream(fileUrl);
            target = new FileOutputStream(outpath);
            Map<String, String> params = new HashMap<>();

            PdfOptions options = PdfOptions.create();

            WordUtil.wordConverterToPdf(source, target, options, params);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
//    @GetMapping(value = "/wordToPdf2")
//    public void buildPdfFromTestWordTemplate() throws IOException
//    {
    //        URL url = this.getClass().getResource("/");
//        String classesUrl = url.getPath();
//        String fileUrl = classesUrl + "template/test.docx";
//        XWPFDocument doc = wordService.createWord();
//
//        //输出渲染后的文件
//        FileOutputStream out = new FileOutputStream("C:\\Users\\xyyto\\Desktop\\output.pdf");
//        PdfOptions options = PdfOptions.create();
//        PdfConverter.getInstance().convert(doc, out, options);
//    }
//
//    @GetMapping(value = "/getWord")
//    public void getWord() throws IOException
//    {
//        XWPFDocument doc = wordService.createWord();
//
//        //输出渲染后的文件
//        FileOutputStream out = new FileOutputStream("C:\\Users\\xyyto\\Desktop\\output.docx");
//
//        doc.write(out);
//        out.flush();
//        out.close();
//    }
}
