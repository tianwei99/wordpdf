package com.test.word.wordpdf.service.impl;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.RenderData;
import com.deepoove.poi.data.TableRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.test.word.wordpdf.service.IWordService;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class WordServiceImpl implements IWordService
{

   /**
    * 生成word
    * @throws IOException
    */
   @Override
   public XWPFDocument createWord()
   {
      ArrayList<RenderData> renderTableHeader = new ArrayList<RenderData>() {{
//         add(new TextRenderData("FFD39B", "疾病"));
         add(new TextRenderData("FFD39B", "药物"));
         add(new TextRenderData("FFD39B", "基因"));
         add(new TextRenderData("FFD39B", "rs"));
         add(new TextRenderData("FFD39B", "证据等级"));
         add(new TextRenderData("FFD39B", "基因型"));
         add(new TextRenderData("FFD39B", "临床指导"));
      }};

      List<Object> renderData = new ArrayList<Object>() {{
         add("懒癌晚期1;懒癌晚期2");
         add("走出小黑屋");
         add("普通的基因");
         add("rs");
         add("S级");
         add("平民型");
         add("再抢救一下");
      }};

      TableRenderData no_datas = new TableRenderData(renderTableHeader, renderData, "no datas", 8600);
      Map<String, Object> datas = new HashMap<String, Object>(){{
         put("name", "老王");
         put("age", "80");
         put("sex","男");
         put("table", no_datas);
      }};

      URL url = this.getClass().getResource("/");
      String classesUrl = url.getPath();
      String templateFileUrl = classesUrl + "template/test.docx";

      //读取模板，进行渲染
      XWPFTemplate template = XWPFTemplate.compile(templateFileUrl).render(datas);
      XWPFDocument xwpfDocument = template.getXWPFDocument();
      return xwpfDocument;
   }

}
