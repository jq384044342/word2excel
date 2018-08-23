package www.qige.service;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.apache.commons.lang.StringUtils;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import sun.misc.BASE64Encoder;
import www.qige.entity.PatEntity;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by Administrator on 2018/6/20 0020.
 */
public class FreeMarker2WordFileUtils {

    /**
     * freemark模板配置
     */
    private  Configuration configuration;
    /**
     * freemark模板的名字
     */
    private  String templateName;
    /**
     * 生成文件名
     */
    private  String fileName;
    /**
     * 生成文件路径
     */
    private  String filePath;

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }

    public String getTemplateName() {
        return templateName;
    }

    public void setTemplateName(String templateName) {
        this.templateName = templateName;
    }
    /**
     * freemark初始化
     * @param templatePath 模板文件位置
     */
    public FreeMarker2WordFileUtils() {

        configuration = new Configuration();
        configuration.setDefaultEncoding("utf-8");
//        String path = FreeMarker2WordFileUtils.class.getClassLoader().getResource("template").getPath();
//        configuration.setClassForTemplateLoading(this.getClass(),path);
//        System.out.println("****The path is ****" + path);
        String path = "template";
        try {
            configuration.setDirectoryForTemplateLoading(new File(path));
        } catch (IOException e) {
            System.out.println("path error");
            e.printStackTrace();
        }
    }
    /**
     * 得到图片
     * @return
     */
    private String getImageStr() {
        String imgFile = "D:/hanmanyi/pic/111.jpg";//需要在D盘下指定的目录下放一张图片
        InputStream in = null;
        byte[] data = null;
        try {
            in = new FileInputStream(imgFile);
            data = new byte[in.available()];
            in.read(data);
            in.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        BASE64Encoder encoder = new BASE64Encoder();//这里会报错
        return encoder.encode(data);
    }


    public  void createWord(PatEntity patEntity) {
        if(patEntity.getProgram().contains("抗AQp4抗体")){
            templateName = "抗AQP4抗体.ftl";
        }else if(patEntity.getProgram().contains("抗GABA B抗体")){
            templateName = "抗GABA B抗体.ftl";
        }else if(patEntity.getProgram().contains("抗GQ1b抗体")){
            templateName = "抗GQ1b抗体.ftl";
        }else if(patEntity.getProgram().contains("抗MBP抗体")){
            templateName = "抗MBP抗体.ftl";
        }else if(patEntity.getProgram().contains("抗MOG抗体")){
            templateName = "抗MOG抗体.ftl";
        }else if(patEntity.getProgram().contains("抗NMDAR抗体")){
            templateName = "抗NMDAR抗体.ftl";
        }else if(patEntity.getProgram().contains("神经节苷脂三项抗体")
                ||patEntity.getProgram().contains("神经节苷脂3项抗体")){
            templateName = "神经节苷脂三项抗体.ftl";
        }else if(patEntity.getProgram().contains("神经节苷脂七项抗体")
                ||patEntity.getProgram().contains("神经节苷脂7项抗体")){
            templateName = "神经节苷脂七项抗体.ftl";
        }else if(patEntity.getProgram().contains("中枢神经脱髓鞘两项抗体")
                ||patEntity.getProgram().contains("中枢神经脱髓鞘2项抗体")){
            templateName = "中枢神经脱髓鞘两项抗体.ftl";
        }else if(patEntity.getProgram().contains("中枢神经脱髓鞘三项抗体")
                ||patEntity.getProgram().contains("中枢神经脱髓鞘3项抗体")){
            templateName = "中枢神经脱髓鞘三项抗体.ftl";
        }else if(patEntity.getProgram().contains("抗神经细胞抗体谱七项抗体")
                ||patEntity.getProgram().contains("抗神经细胞抗体谱7项抗体")){
            templateName = "抗神经细胞抗体谱7项抗体.ftl";
        }else if(patEntity.getProgram().contains("抗神经细胞抗体谱九项抗体")
                ||patEntity.getProgram().contains("抗神经细胞抗体谱9项抗体")){
            templateName = "抗神经细胞抗体谱9项抗体.ftl";
        }else if(patEntity.getProgram().contains("抗神经细胞抗体谱八项抗体")
                ||patEntity.getProgram().contains("抗神经细胞抗体谱9项抗体")){
            templateName = "抗神经细胞抗体谱8项抗体.ftl";
        }else if(patEntity.getProgram().contains("抗神经细胞抗体谱六项抗体")
                ||patEntity.getProgram().contains("抗神经细胞抗体谱6项抗体")){
            templateName = "抗神经细胞抗体谱6项抗体.ftl";
        }else if(patEntity.getProgram().contains("抗GFAP抗体")){
            templateName = "抗GFAP抗体.ftl";
        }else if(patEntity.getProgram().contains("抗NF155抗体")){
            templateName = "抗NF155抗体.ftl";
        }else if(patEntity.getProgram().contains("抗NF186抗体")){
            templateName = "抗NF186抗体.ftl";
        }else if(patEntity.getProgram().contains("郎飞氏结相关2项抗体")
                ||patEntity.getProgram().contains("郎飞氏结相关二项抗体")||patEntity.getProgram().contains("郎飞氏结相关两项抗体")){
            templateName = "郎飞氏结相关2项抗体.ftl";
        }else {
            return;
        }
//        抗AQP4抗体检测报告.余飞荣.女.29.血清.2018.07.2.doc
        String[] testDate = patEntity.getTestDate().split("年");
        String year = testDate[0];
        String[] months = testDate[1].split("月");
        String month = Integer.valueOf(months[0])<10?0+months[0]:months[0];
        String day = Integer.valueOf(months[1].split("日")[0])<10?0+months[1].split("日")[0]:months[1].split("日")[0];
        fileName = patEntity.getProgram()+"报告."+patEntity.getName()+"."+patEntity.getSex()+"."+patEntity.getAge()
                +"."+patEntity.getType()+"."+year+"."+month+"."+day+".doc";//生成的word名称
//        filePath = "D:\\temp1\\temp\\";//生成word存储路径
//        String filePath1 = "D:/temp2/";//生成word存储路径
//
        filePath = "report\\temp\\";//生成word存储路径
        String filePath1 = "report/";
        Template t = null;
        try {
            t = configuration.getTemplate(templateName);
        } catch (IOException e) {
            e.printStackTrace();
        }

        File outFile = new File(filePath + fileName);
        System.out.println(outFile.getAbsolutePath());
        //获取父目录
        File fileParent = outFile.getParentFile();
        //判断是否存在
        if (!fileParent.exists()) {
            //创建父目录文件
            fileParent.mkdirs();
        }
        Writer out = null;
        try {
            out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile), "UTF-8"));
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        Map map = new HashMap<String, Object>();

        map.put("name", StringUtils.isEmpty(patEntity.getName())?" ":patEntity.getName());
        map.put("type", StringUtils.isEmpty(patEntity.getType())?" ":patEntity.getType());
        map.put("sex", StringUtils.isEmpty(patEntity.getSex())?" ":patEntity.getSex());
        map.put("age", StringUtils.isEmpty(patEntity.getAge())?" ":patEntity.getAge());
        map.put("start", StringUtils.isEmpty(patEntity.getStart())?" ":patEntity.getStart());
        map.put("collectTime", StringUtils.isEmpty(patEntity.getCollectTime())?" ":patEntity.getCollectTime());
//        map.put("testPeo",StringUtils.isEmpty(patEntity.getTestPeo())?" ": patEntity.getTestPeo());
//        map.put("reporter", StringUtils.isEmpty(patEntity.getReporter())?" ":patEntity.getReporter());
//        map.put("audit", StringUtils.isEmpty(patEntity.getAudit())?" ":patEntity.getAudit());
        map.put("testDate", StringUtils.isEmpty(patEntity.getTestDate())?" ":patEntity.getTestDate());
        map.put("reporterDate",StringUtils.isEmpty(patEntity.getReporterDate())?" ": patEntity.getReporterDate());

        try {
            t.process(map, out);
            out.close();
        } catch (TemplateException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        ActiveXComponent _app = new ActiveXComponent("Word.Application");
        _app.setProperty("Visible", Variant.VT_FALSE);

        Dispatch documents = _app.getProperty("Documents").toDispatch();

        // 打开FreeMarker生成的Word文档
        String abPath = outFile.getAbsolutePath().replace(fileName,"");
        System.out.println("=======" + abPath+fileName);
        Dispatch doc = Dispatch.call(documents, "Open", abPath+fileName, Variant.VT_FALSE, Variant.VT_FALSE).toDispatch();
        // 另存为新的Word文档
        System.out.println(abPath.substring(0,abPath.length()-5)+fileName);
        Dispatch.call(doc, "SaveAs", abPath.substring(0,abPath.length()-5)+fileName, Variant.VT_FALSE, Variant.VT_FALSE);
        Dispatch.call(doc, "Close", Variant.VT_FALSE);
        _app.invoke("Quit", new Variant[] {});
        ComThread.Release();
        outFile.delete();
    }

    public static void main(String[] args){
        String a = "1";
        System.out.println(a+"/");
    }

}
