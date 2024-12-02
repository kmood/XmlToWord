package com.test;


import com.kmood.datahandle.DocumentProducer;
import com.kmood.utils.ConvertionUtil;
import com.kmood.utils.FileUtils;
import com.kmood.utils.FreemarkerUtil;
import com.kmood.utils.JsonBinder;
import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import com.kmood.word.WordModelParser;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.junit.Test;

import java.io.*;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @Auther: SunBC
 * @Date: 2019/1/15 12:29
 * @Description:  TODO 待解决 1.xml的模板导出目前出错，仅支持docx的模板 2.docx生成的临时文件目录删除失败，不影响报告生成。
 */
public class test {


    /**
     * description:包装说明表（范例A）.xml  模板导出测试,验证格式，测试转义字符。
     * @auther: SunBC
     * @date: 2019/7/12 16:58
     */
    @Test
    public  void test1() throws IOException, TemplateException {
        try {
            String ActualModelPath = this.getClass().getClassLoader().getResource("./model/").toURI().getPath();
            String xmlPath = this.getClass().getClassLoader().getResource("./model").toURI().getPath();
            String ExportFilePath = this.getClass().getClassLoader().getResource("./export").toURI().getPath() + "/包装说明表（范例A）.doc";
            HashMap<String, Object> map = new HashMap<>();
            map.put("zzdhm", "kmood-制造单号码");
            map.put("ydwcrq", "kmood-预定完成日期");
            map.put("cpmc", "kmood-产品名称");
            map.put("jyrq", "kmood-交运日期");
            map.put("sl", "kmood-数量");
            map.put("xs", "kmood-箱数");

            ArrayList<Object> zxsmList = new ArrayList<>();
            HashMap<String, Object> zxsmmap = new HashMap<>();
            zxsmmap.put("xh", "kmood-箱号");
            zxsmmap.put("xs", "kmood-箱数");
            zxsmmap.put("zrl", "kmood-梅香");
            zxsmmap.put("zsl", "kmood-交运日期");
            zxsmmap.put("sm", "kmood-交运日期");
            zxsmList.add(zxsmmap);
            HashMap<String, Object> zxsmmap1 = new HashMap<>();
            zxsmmap1.put("xh", "kmood-制造单号码");
            zxsmmap1.put("xs", "kmood-预定完成日期");
            zxsmmap1.put("zrl","kmood-产品名称");
            zxsmmap1.put("zsl","kmood-交运日期");
            zxsmmap1.put("sm", "kmood-交运日期");
            zxsmList.add(zxsmmap);
            map.put("zxsm", zxsmList);
            map.put("sbsm", "kmood-商标说明");
            map.put("bt", "kmood OfficeExport 导出word");
            URL titleUrl = this.getClass().getClassLoader().getResource("./picture/exportTestPicture-title.png");
            String intro = Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath(titleUrl.toURI().getPath()));

            map.put("mypicture",intro);
            DocumentProducer dp = new DocumentProducer(ActualModelPath);
            String complie = dp.Complie(xmlPath, "包装说明表（范例A）.xml", true);
            System.out.println(complie);
            dp.produce(map, ExportFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * description:测试导出图片
     * @auther: SunBC
     * @date: 2019/7/16 17:29
     */
    @Test
    public void testexportPicture () {
        try {
            URL Url = this.getClass().getClassLoader().getResource("./model/picture.xml");
            HashMap<String, Object> map = new HashMap<>();
            URL introUrl = this.getClass().getClassLoader().getResource("./picture/exportTestPicture-intro.png");
            URL codeUrl = this.getClass().getClassLoader().getResource("./picture/exportTestPicture-code.png");
            URL titleUrl = this.getClass().getClassLoader().getResource("./picture/exportTestPicture-title.png");
            String intro = Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath(introUrl.toURI().getPath()));

            map.put("intro", intro);
            String code = Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath(codeUrl.toURI().getPath()));
            map.put("code", code);
            map.put("title", Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath(titleUrl.toURI().getPath())));
            String path = Url.toURI().getPath();
            String exportPath = path + ".doc";
            String compile = WordModelParser.Compile(path, null);
            OutputStreamWriter outputStreamWriter = new OutputStreamWriter(new FileOutputStream(exportPath), "utf-8");
            Template template = FreemarkerUtil.configuration.getTemplate("picture.xml.ftl");
            template.process(map, outputStreamWriter);
            System.out.println("-----导出文件路径-----" + exportPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void exportDb()throws Exception{

        DocumentProducer dp = new DocumentProducer("D:\\intelliJ IDEA_workerspace\\ngccoa\\src\\main\\resources\\model");
        String complie = dp.Complie("D:\\intelliJ IDEA_workerspace\\ngccoa\\src\\main\\resources\\model", "fwngnew.xml", true);

    }
    /**
     * 测试2003 版导出图片 doc文档
     * @throws Exception
     */
    @Test
    public void testpic() throws Exception {
        ClassLoader classLoader = this.getClass().getClassLoader();
        if (classLoader == null){
            classLoader = ClassLoader.getSystemClassLoader();
        }
        String ActualModelPath = classLoader.getResource("./model/").toURI().getPath();
        String xmlPath = classLoader.getResource("./model").toURI().getPath();
        String ExportFilePath = classLoader.getResource("./export").toURI().getPath() + "/picture.doc";

        HashMap<String, Object> map = new HashMap<>();
        //读取输出图片
        URL introUrl = classLoader.getResource("./picture/exportTestPicture-intro.png");
        URL codeUrl = classLoader.getResource("./picture/exportTestPicture-code.png");
        URL titleUrl = classLoader.getResource("./picture/exportTestPicture-title.png");

        String intro = Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath(introUrl.toURI().getPath()));
        map.put("intro", intro);
        String code = Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath(codeUrl.toURI().getPath()));
        map.put("code", code);
        map.put("title", Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath(titleUrl.toURI().getPath())));
        //编译输出
        DocumentProducer dp = new DocumentProducer(ActualModelPath);
        String complie = dp.Complie(xmlPath, "picture.xml", true);
        dp.produce(map, ExportFilePath);
    }

    /**
     * 测试docx 模板导出
     * @throws IOException
     * @throws TemplateException
     */
    @Test
    public void testdocx() throws IOException, TemplateException {
        try {
            String ActualModelPath = this.getClass().getClassLoader().getResource("./model/").toURI().getPath();
            String xmlPath = this.getClass().getClassLoader().getResource("./model").toURI().getPath();
            String filePath = this.getClass().getClassLoader().getResource("./picture/exportTestPicture-code.png").toURI().getPath();
            String filePath2 = this.getClass().getClassLoader().getResource("./picture/exportTestPicture-intro.png").toURI().getPath();
            String filePath3 = this.getClass().getClassLoader().getResource("./picture/exportTestPicture-title.png").toURI().getPath();
            String ExportFilePath = this.getClass().getClassLoader().getResource("./export").toURI().getPath() + "/包装说明表（范例A）-export.docx";
            HashMap<String, Object> map = new HashMap<>();
            map.put("zzdhm", "yangzh-\n制造单\n测试内容测试内容测试内容测试内容测试内容测试内容测试内容测试内容测试内容测试内容");
            map.put("ydwcrq", "2024-12-30");
            map.put("cpmc", "产品名称");
            map.put("jyrq", "2024-12-02");
            map.put("sl", "200");
            map.put("xs", "100");

            ArrayList<Object> zxsmList = new ArrayList<>();
            HashMap<String, Object> zxsmmap = new HashMap<>();
            zxsmmap.put("xh", "01");
            zxsmmap.put("xs", "10\n（箱数²）");
            zxsmmap.put("zrl", "100");
            zxsmmap.put("zsl", "50");
            zxsmmap.put("sm", "2024-12-02");
            // 图片方式一（推荐）：图片的绝对路径的地址
            Map pictestMap=new HashMap();
            pictestMap.put("type","pic");
            pictestMap.put("url",filePath3);
            Map pictext_loopMap=new HashMap();
            pictext_loopMap.put("type","pic");
            pictext_loopMap.put("url",filePath3);
            zxsmmap.put("pictest", pictestMap);
            zxsmmap.put("pictext_loop", pictext_loopMap);

            zxsmList.add(zxsmmap);
            HashMap<String, Object> zxsmmap1 = new HashMap<>();
            zxsmmap1.put("xh", "02");
            zxsmmap1.put("xs", "kmood-预定完成日期");
            zxsmmap1.put("zrl","kmood-产品名称");
            zxsmmap1.put("zsl","kmood-交运日期");
            zxsmmap1.put("sm", "kmood-交运日期");
            // 图片方式二（不推荐）：base64字符串，图片过大可能会失败
            zxsmmap1.put("pictest", Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath( filePath)));
            zxsmmap1.put("pictext_loop", Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath( filePath)));

            zxsmList.add(zxsmmap1);
            map.put("zxsm", zxsmList);
            map.put("sbsm", "yangzh-商标说明");
            map.put("mp", "kmood OfficeExport 导出word");
            map.put("pictext",  Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath( filePath2)));

            DocumentProducer dp = new DocumentProducer(ActualModelPath);
            String complie = dp.Complie(xmlPath, "包装说明表（范例A）.docx", false);
            System.out.println(complie);
            dp.produce(map, ExportFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }




    /**
     * 测试 图片编码
     * @throws Exception
     */
    @Test
    public void endPic() throws Exception {

        String intro = Base64.getEncoder().encodeToString(FileUtils.readToBytesByFilepath("E:\\project\\gitlab\\qgzhdc\\qgzhdc-towordpdf-java\\test.png"));
        System.out.println(intro);
    }

}
