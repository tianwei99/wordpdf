package com.test.word.wordpdf.controller;

import com.test.word.wordpdf.service.IWordService;
import com.test.word.wordpdf.util.WordUtil;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;

@RestController
public class WordController
{
    @Autowired
    private IWordService wordService;

	@GetMapping(value = "/wordToPdf")
	public void wordToPdf()
	{
        URL url = this.getClass().getResource("/");
        String classesUrl = url.getPath();
        String fileUrl = classesUrl + "template/test.docx";

        String outpath = "C:\\Users\\xyyto\\Desktop\\output.pdf";

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

    @GetMapping(value = "/wordToPdf2")
    public void buildTestWordTemplate() throws IOException
    {
        XWPFDocument doc = wordService.createWord();

        //输出渲染后的文件
        FileOutputStream out = new FileOutputStream("C:\\Users\\xyyto\\Desktop\\output.pdf");
        PdfOptions options = PdfOptions.create();
        PdfConverter.getInstance().convert(doc, out, options);

        doc.write(out);
        out.flush();
        out.close();
    }
}
