package com.kmood.datahandle;

import com.kmood.utils.FileUtils;
import com.kmood.utils.StringUtil;
import com.kmood.utils.ZipUtils;
import com.kmood.utils.dom4jUtils;
import com.kmood.word.WordModelHandlerImpl;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import net.lingala.zip4j.ZipFile;
import org.apache.commons.lang3.StringUtils;
import org.dom4j.*;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

import java.io.*;
import java.util.*;

public class DocumentProducer {
    private ThreadLocal<String> ActualModelPathLocal = new ThreadLocal<>();
    private ThreadLocal<String> ActualModelNameLocal = new ThreadLocal<>();
    private ThreadLocal<String> ActualModelOriginPathLocal = new ThreadLocal<>(); // 记录模板文件绝对路径
    private ThreadLocal<String> ActualModelDocxPathLocal = new ThreadLocal<>(); // 记录docx文件解压后的xml文件路径
    private ThreadLocal<String> ActualExtractDocxPathLocal = new ThreadLocal<>(); // 记录通过模板解压后文件路径
    public static ThreadLocal<String> ModelSuffixFlagLocal = new ThreadLocal<>(); // 记录当前模板类型[xml,docx]

    public static ThreadLocal<Integer> PicCountLocal=new ThreadLocal<>();
    public DocumentProducer(Configuration CustomerConfig,String ActualModelPath){
        FMConfiguration.Init(CustomerConfig);
        ActualModelPathLocal.set(ActualModelPath);
    }
    public DocumentProducer(String ActualModelPath)throws Exception{
        FMConfiguration.Init();
//        Configuration configuration = FMConfiguration.addFMModelPath(ActualModelPath);
        ActualModelPathLocal.set(ActualModelPath);
    }
    public DocumentProducer()throws Exception{
        FMConfiguration.Init();
    }
    public String Complie(String XmlModelPath,String XmlModelName,boolean debugModel)throws Exception {
        // 增加支持docx文件模板渲染
        ModelSuffixFlagLocal.set("xml");
        boolean isDocx = "docx".equalsIgnoreCase(FileUtils.getFileSuffixByPath(XmlModelPath + File.separator + XmlModelName));
        if(isDocx){
            // 如果是docx文件，先解压到临时目录
            String uid= UUID.randomUUID().toString();
            ModelSuffixFlagLocal.set("docx");
            ActualModelOriginPathLocal.set(XmlModelPath+File.separator+XmlModelName);//  eg: G:/qgzhdc/officeexport-java/target/test-classes/model\包装说明表（范例A）.docx
            String docxFilePath=XmlModelPath+File.separator+XmlModelName;
            String compressPath=XmlModelPath+File.separator+XmlModelName.substring(0,XmlModelName.lastIndexOf(".") )+"_"+uid; // G:/qgzhdc/officeexport-java/target/test-classes/model\包装说明表（范例A）
            ActualExtractDocxPathLocal.set(compressPath);
           // new ZipFile(docxFilePath).extractAll(compressPath);
            new ZipFile(docxFilePath).extractAll(compressPath);
            XmlModelPath=XmlModelPath+File.separator+XmlModelName.substring(0,XmlModelName.lastIndexOf(".")) +"_"+uid+File.separator+"word"+File.separator;
            XmlModelName="document.xml";
            FMConfiguration.clearFMModelPathArr();
            ActualModelPathLocal.remove();
            ActualModelDocxPathLocal.set(XmlModelPath+File.separator+XmlModelName);
        }
        if(StringUtil.isBlank(ActualModelPathLocal.get())){
            FMConfiguration.addFMModelPath(XmlModelPath);
            ActualModelPathLocal.set(XmlModelPath);
        }
        if (debugModel || isDocx){
            WordModelHandlerImpl wordXmlModelHandler = new WordModelHandlerImpl();
            String path = wordXmlModelHandler.WordXmlModelHandle(XmlModelPath+File.separator+XmlModelName,ActualModelPathLocal.get());
        }
        ActualModelNameLocal.set(XmlModelName+".ftl");
        return XmlModelPath+File.separator+XmlModelName+".ftl";
    }
    //增加构造函数
    public void produce(Object data,String ProduceFilePath) throws Exception {
        Writer outputStreamWriter = null;

        try{
        String produceFilePathOrigin=ProduceFilePath;
        if("docx".equalsIgnoreCase(ModelSuffixFlagLocal.get())  ){
            ProduceFilePath=ActualModelDocxPathLocal.get();
        }
        Configuration configuration = FMConfiguration.getConfiguration();
        String modelFileName=ActualModelNameLocal.get();
        String modelPath=ActualModelPathLocal.get();
        Template template = configuration.getTemplate(modelFileName);
        outputStreamWriter = new OutputStreamWriter(new FileOutputStream(ProduceFilePath), template.getEncoding());
        // 如果是docx,模拟图片插入
        if("docx".equalsIgnoreCase(ModelSuffixFlagLocal.get())  ){
            data = DataConverter.addPictureXh(data,"");
            File metaFolder=new File(modelPath+File.separator+"media");
            if(!metaFolder.exists()){
                metaFolder.mkdir();
            }
            HashMap renderData=(HashMap) data;
            dealPicture(renderData,modelPath);
        }

        Object dataConvert = DataConverter.convert(data, null);
        template.process(dataConvert,outputStreamWriter);
        outputStreamWriter.flush();
        outputStreamWriter.close();
        //处理换行

        SAXReader reader = new SAXReader();
        File file = new File(ProduceFilePath);
        Document document = reader.read(file);

        List ParagList = document.selectNodes(".//w:p");

        for (int i = 0; i < ParagList.size(); i++) {
            Element WPNode = (Element) ParagList.get(i);
            List WPChildList = WPNode.elements();
            for (int j = 0; j < WPChildList.size(); j++) {
                Element childEle = (Element)WPChildList.get(j);
                Element ele = null;
                String name = childEle.getName();
                if(StringUtils.equals(name,"r")){
                    List<Element> wtList = childEle.selectNodes("w:t");
                    for (Element wtEle:wtList) {
                        String text = wtEle.getText();
                        if(text.contains("yangzh-制造单")){
                            System.out.println("");
                        }
                        if(text.contains("\n")){
                            String[] strArr = text.split("\n");
                            if(strArr.length>1){
                                for (int k = 0; k < strArr.length; k++) {
                                    String s = strArr[k];
                                    if(k==0){
                                        wtEle.setText(s);
                                        ele = wtEle;
                                    }else{
                                        Element wbrEle = DocumentHelper.createElement("w:br");
                                        wbrEle.addAttribute("w:clear","right");
                                        dom4jUtils.addSiblingElement(ele, wbrEle);
                                        Element copyWtEle = wtEle.createCopy();
                                        copyWtEle.setText(s);
                                        dom4jUtils.addSiblingElement(wbrEle, copyWtEle);
                                        ele = copyWtEle;
                                    }

                                }
                            }

                        }
                    }
                }
            }
        }
        outputStreamWriter = new OutputStreamWriter(new FileOutputStream(ProduceFilePath,false), template.getEncoding());
        document.write(outputStreamWriter);
        outputStreamWriter.flush();
        if("docx".equalsIgnoreCase(ModelSuffixFlagLocal.get())  ){
            String sourceFolderPath=ActualExtractDocxPathLocal.get();
            new File(ProduceFilePath+".ftl").delete();
            ZipUtils.compress(produceFilePathOrigin,sourceFolderPath);
            // 删除生成的临时文件
            System.out.println(FileUtils.deleteDir(new File(sourceFolderPath))) ;
            FMConfiguration.clearFMModelPathArr();
        }
        }finally {
            if (outputStreamWriter != null )
                outputStreamWriter.close();
        }
    }
    public void produce(Object data,OutputStream ProduceFileout)throws IOException,TemplateException {
        Configuration configuration = FMConfiguration.getConfiguration();
        Template template = configuration.getTemplate(ActualModelNameLocal.get());
        OutputStreamWriter outputStreamWriter = new OutputStreamWriter(ProduceFileout, template.getEncoding());
        Object dataConvert = DataConverter.convert(data, null);
        template.process(dataConvert,outputStreamWriter);
    }
    public void produce(Object data,OutputStream ProduceFileout,Config config)throws IOException,TemplateException{
        Configuration configuration = FMConfiguration.getConfiguration();
        Template template = configuration.getTemplate(ActualModelNameLocal.get());
        OutputStreamWriter outputStreamWriter = new OutputStreamWriter(ProduceFileout, template.getEncoding());
        Object dataConvert = DataConverter.convert(data, config);
        template.process(dataConvert,outputStreamWriter);
    }
    public void produce(Object data,String ProduceFilePath,Config config)throws IOException,TemplateException{
        Configuration configuration = FMConfiguration.getConfiguration();
        Template template = configuration.getTemplate(ActualModelNameLocal.get());
        OutputStreamWriter outputStreamWriter = new OutputStreamWriter(new FileOutputStream(ProduceFilePath), template.getEncoding());
        Object dataConvert = DataConverter.convert(data, config);
        template.process(dataConvert,outputStreamWriter);
    }

    /**
     *
     * @param renderData
     * @param modelPath
     * @throws Exception
     */
    private void dealPicture(Map renderData,String modelPath) throws Exception {
        for(Object key : renderData.keySet()){ // 这块应该使用递归，获取所有的图片并输出
            Object obj= renderData.get(key);
            if(obj!=null&&obj.getClass().equals(ArrayList.class)){
                ArrayList mapList=(ArrayList)obj;
                for(Object mapItem :mapList){
                    dealPicture((Map)mapItem,modelPath);
                }
            }
//            String base64png=renderData.get(key)!=null? renderData.get(key).toString():"";
            String base64png= FileUtils.convertParamsToBase64(renderData.get(key));
            String bas64flag=FileUtils.checkImageBase64Format(base64png);
            if("png".equalsIgnoreCase(bas64flag) || "jpg".equalsIgnoreCase(bas64flag) ||"jpeg".equalsIgnoreCase(bas64flag)){
                    if("jpg".equalsIgnoreCase(bas64flag)){
                        bas64flag = "jpeg";
                    }
                    // 新建图片
                    if(renderData.containsKey("_xh")){
                        FileUtils.convertBase64DataToImage(base64png,modelPath+File.separator+"media"+File.separator+"image"+renderData.get("_xh")+key+"."+bas64flag);
                    }else{
                        FileUtils.convertBase64DataToImage(base64png,modelPath+File.separator+"media"+File.separator+"image"+key+"."+bas64flag);
                    }

                // 在docx文件中增加png的引用
                SAXReader reader = new SAXReader();
                File relsDocumentXmlFile = new File(modelPath+File.separator+"_rels"+File.separator+"document.xml.rels");
                Document document = reader.read(relsDocumentXmlFile);
                Element element = document.getRootElement();
                Element pictureRelationship = element.addElement("Relationship");

                pictureRelationship.addAttribute("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                //http://schemas.openxmlformats.org/officeDocument/2006/relationships/image

                if(renderData.containsKey("_xh")){
                    pictureRelationship.addAttribute("Id","rId"+renderData.get("_xh")+key);
                    pictureRelationship.addAttribute("Target","media/image"+renderData.get("_xh")+key+"."+bas64flag);
                }else{
                    pictureRelationship.addAttribute("Id","rId"+key);
                    pictureRelationship.addAttribute("Target","media/image"+key+"."+bas64flag);
                }

                FileOutputStream fos=new FileOutputStream(modelPath+File.separator+"_rels"+File.separator+"document.xml.rels");
                XMLWriter writer = new XMLWriter(fos); //创建xml输出流
                writer.write(document);
                writer.close();

                // 存在图片占位  [Content_Types].xml增加  png引用。
                File contentTypesFolderFile = new File(modelPath).getParentFile();
                File contentTypesXmlFile=new File(contentTypesFolderFile.getAbsolutePath()+File.separator+"[Content_Types].xml");
                SAXReader readernew = new SAXReader();
                Document documentnew = readernew.read(contentTypesXmlFile);
                Element rootElementnew = documentnew.getRootElement();
                List<Element> elementsnew = rootElementnew.elements();
                List<Element> elementDeaultList = rootElementnew.elements("Default");
                boolean sfczpng=false;
                boolean sfczjpeg=false;
                for(Element elementitem:elementDeaultList){
                    Attribute extension = elementitem.attribute("Extension");
                    // System.out.println(extension.getData());
                    if("png".equalsIgnoreCase(extension.getData().toString())){
                       sfczpng=true;
                    }
                    if("jpeg".equalsIgnoreCase(extension.getData().toString())){
                        sfczjpeg=true;
                    }
                }
                if(!sfczpng&&"png".equalsIgnoreCase(bas64flag)){
                    // docx中在[Content_types].xml文件增加 png配置，并且需要将Default节点放到前面，否则docx格式不正确
                    Element pictureRelationshipnew =  DocumentHelper.createElement(QName.get("Default", rootElementnew.getNamespace ()));
                    pictureRelationshipnew.addAttribute("Extension","png");
                    pictureRelationshipnew.addAttribute("ContentType","image/png");
                    elementsnew.add(0,pictureRelationshipnew);
                    FileOutputStream fosnew=new FileOutputStream(contentTypesFolderFile.getAbsolutePath()+File.separator+"[Content_Types].xml");
                    XMLWriter writernew = new XMLWriter(fosnew); //创建xml输出流
                    writernew.write(documentnew);
                    writernew.close();
                }
                if(!sfczjpeg&&"jpeg".equalsIgnoreCase(bas64flag)){
                    // docx中在[Content_types].xml文件增加 png配置，并且需要将Default节点放到前面，否则docx格式不正确
                    Element pictureRelationshipnew =  DocumentHelper.createElement(QName.get("Default", rootElementnew.getNamespace ()));
                    pictureRelationshipnew.addAttribute("Extension","jpeg");
                    pictureRelationshipnew.addAttribute("ContentType","image/jpeg");
                    elementsnew.add(0,pictureRelationshipnew);
                    FileOutputStream fosnew=new FileOutputStream(contentTypesFolderFile.getAbsolutePath()+File.separator+"[Content_Types].xml");
                    XMLWriter writernew = new XMLWriter(fosnew); //创建xml输出流
                    writernew.write(documentnew);
                    writernew.close();
                }

            }
        }
    }



}
