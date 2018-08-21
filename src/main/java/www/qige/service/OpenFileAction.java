package www.qige.service;

import org.apache.commons.lang.StringUtils;
import sun.awt.windows.ThemeReader;
import www.qige.entity.PatEntity;


import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class OpenFileAction extends JFrame implements ActionListener{
    JButton jb = new JButton("打开文件");

    public OpenFileAction(){
        jb.setActionCommand("open");
        jb.setBackground(Color.CYAN);//设置按钮颜色
        this.getContentPane().add(jb, BorderLayout.SOUTH);//建立容器使用边界布局
        //  
        jb.addActionListener(this);
        this.setTitle("Excel生成Word小工具");
        this.setSize(333, 288);
        this.setLocation(200,200);
        //显示窗口true  
        this.setVisible(true);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    }
    public void actionPerformed(ActionEvent e){
        if (e.getActionCommand().equals("open")){
            JFileChooser jf = new JFileChooser();
            jf.showOpenDialog(this);//显示打开的文件对话框  
            File f =  jf.getSelectedFile();//使用文件类获取选择器选择的文件  
            String path = f.getAbsolutePath();//返回路径名
            ReadExcelUtils readExcelUtils = new ReadExcelUtils(path);
            java.util.List<PatEntity> list = new ArrayList<PatEntity>();
            Long start = System.currentTimeMillis();
            try {
                Map<Integer, Map<Integer, Object>> integerMapMap = readExcelUtils.readExcelContent();
                for(Map.Entry<Integer, Map<Integer, Object>> entry : integerMapMap.entrySet()){
                    if(StringUtils.isEmpty((String) entry.getValue().get(0))){
                        continue;
                    }
                    PatEntity pat = new PatEntity();
                    pat.setName((String)entry.getValue().get(2));
                    pat.setType((String)entry.getValue().get(7));
                    pat.setSex((String)entry.getValue().get(3));
                    String age = (String) entry.getValue().get(4);
                    if(age.contains(".0")){
                        age = age.replace(".0","");
                    }
                    pat.setAge(age);
                    pat.setStart((String)entry.getValue().get(9));
                    pat.setCollectTime((String)entry.getValue().get(9));
//                    pat.setTestPeo((String)entry.getValue().get(15));
//                    pat.setReporter((String)entry.getValue().get(11));
//                    pat.setAudit((String)entry.getValue().get(15));
                    pat.setTestDate((String)entry.getValue().get(10));
                    pat.setReporterDate((String)entry.getValue().get(10));

                    pat.setProgram((String)entry.getValue().get(6));
                    list.add(pat);
                }
                System.out.println(list.size());
                ExecutorService cachedThreadPool = Executors.newSingleThreadExecutor();;
                for(PatEntity entity:list){
//                    FreeMarker2WordFileUtils freemark = new FreeMarker2WordFileUtils();
//                    freemark.createWord(entity);//生成方法
                    WordRunable wordRunable = new WordRunable(entity);
                    cachedThreadPool.execute(wordRunable);
                }
                cachedThreadPool.shutdown();
                while (true){
                    if(cachedThreadPool.isTerminated()){
                        long timeused = System.currentTimeMillis()-start;
                        System.out.println("生成报告耗费时间"+timeused+"ms");
                        break;
                    }
                }
                JOptionPane.showMessageDialog(this, "已经生成报告", "标题",JOptionPane.WARNING_MESSAGE);
            } catch (Exception e1) {
                JOptionPane.showMessageDialog(this, "报告生成错误", "标题",JOptionPane.WARNING_MESSAGE);
                e1.printStackTrace();
            }
            //JOptionPane弹出对话框类，显示绝对路径名  
//            JOptionPane.showMessageDialog(this, s, "标题",JOptionPane.WARNING_MESSAGE);
        }
    }
    class WordRunable implements Runnable{
        private PatEntity entity;

        public WordRunable(PatEntity entity) {
            this.entity = entity;
        }

        @Override
        public void run() {
            FreeMarker2WordFileUtils freemark = new FreeMarker2WordFileUtils();
            freemark.createWord(entity);
        }
    }

    public static void main(String[] args) {
        // TODO 自动生成的方法存根
        new OpenFileAction();
    }
}  