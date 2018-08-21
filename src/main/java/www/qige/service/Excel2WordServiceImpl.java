package www.qige.service;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;

/**
 * Created by Administrator on 2018/6/19 0019.
 */
public class Excel2WordServiceImpl {
    public void fromExcel2WordFiles(String FilePath){
        //生成word表格
        try {
            //Blank Document
            XWPFDocument document= new XWPFDocument();

            //Write the Document in file system
            FileOutputStream out = new FileOutputStream(new File("create_table.docx"));


            //添加标题
            XWPFParagraph titleParagraph = document.createParagraph();
            //设置段落居中
            titleParagraph.setAlignment(ParagraphAlignment.CENTER);

            XWPFRun titleParagraphRun = titleParagraph.createRun();
            titleParagraphRun.setText("检验报告单");
            titleParagraphRun.setColor("000000");
            titleParagraphRun.setFontSize(20);


            //段落
//            XWPFParagraph firstParagraph = document.createParagraph();
//            XWPFRun run = firstParagraph.createRun();
//            run.setText("Java POI 生成word文件。");
//            run.setColor("696969");
//            run.setFontSize(16);

            //设置段落背景颜色
//            CTShd cTShd = run.getCTR().addNewRPr().addNewShd();
//            cTShd.setVal(STShd.CLEAR);
//            cTShd.setFill("97FFFF");


            //换行
            XWPFParagraph paragraph1 = document.createParagraph();
            XWPFRun paragraphRun1 = paragraph1.createRun();
            paragraphRun1.setText("\r");


            //基本信息表格
            XWPFTable infoTable = document.createTable();
            //去表格边框
            infoTable.getCTTbl().getTblPr().unsetTblBorders();


            //列宽自动分割
            CTTblWidth infoTableWidth = infoTable.getCTTbl().addNewTblPr().addNewTblW();
            infoTableWidth.setType(STTblWidth.DXA);
            infoTableWidth.setW(BigInteger.valueOf(9072));


            //表格第一行
            XWPFTableRow infoTableRowOne = infoTable.getRow(0);
            infoTableRowOne.getCell(0).setText("姓    名：");
            infoTableRowOne.addNewTableCell().setText("章祝发");
            infoTableRowOne.addNewTableCell().setText("标本类型：");
            infoTableRowOne.addNewTableCell().setText("脑脊液");
            infoTableRowOne.addNewTableCell().setText("标本性状：");
            infoTableRowOne.addNewTableCell().setText("     ");
            infoTableRowOne.addNewTableCell().setText("送检单位：");
            infoTableRowOne.addNewTableCell().setText("     ");

            //表格第二行
            XWPFTableRow infoTableRowTwo = infoTable.createRow();
            infoTableRowTwo.getCell(0).setText("住 院 号：");
            infoTableRowTwo.getCell(1).setText("     ");
            infoTableRowTwo.getCell(2).setText("性    别：");
            infoTableRowTwo.getCell(3).setText("男");
            infoTableRowTwo.getCell(4).setText("年    龄：");
            infoTableRowTwo.getCell(5).setText("76");
            infoTableRowTwo.getCell(6).setText("送检医师：");
            infoTableRowTwo.getCell(7).setText("     ");

            //表格第三行
            XWPFTableRow infoTableRowThree = infoTable.createRow();
            infoTableRowThree.getCell(0).setText("床    号：");
            infoTableRowThree.getCell(1).setText("     ");
            infoTableRowThree.getCell(2).setText("科    别：");
            infoTableRowThree.getCell(3).setText("     ");
            infoTableRowThree.getCell(4).setText("接收时间：");
            infoTableRowThree.getCell(5).setText("2018/6/12");
            infoTableRowThree.getCell(6).setText("采集时间：");
            infoTableRowThree.getCell(7).setText("2018/6/12");

            //表格第三行
            XWPFTableRow infoTableRowFour = infoTable.createRow();
            infoTableRowFour.getCell(0).setText("临床诊断：");
            infoTableRowFour.getCell(2).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            for(int i=3;i<8;i++){
                infoTableRowFour.getCell(i).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
//            infoTableRowFour.getCell(1).setText("     ");
//            infoTableRowFour.getCell(2).setText("     ");
//            infoTableRowFour.getCell(3).setText("     ");
//            infoTableRowFour.getCell(4).setText("     ");
//            infoTableRowFour.getCell(5).setText("     ");
//            infoTableRowFour.getCell(6).setText("     ");
//            infoTableRowFour.getCell(7).setText("     ");
//            //表格第四行
//            XWPFTableRow infoTableRowFour = infoTable.createRow();
//            infoTableRowFour.getCell(0).setText("性别");
//            infoTableRowFour.getCell(1).setText(": 男");
//
//            //表格第五行
//            XWPFTableRow infoTableRowFive = infoTable.createRow();
//            infoTableRowFive.getCell(0).setText("现居地");
//            infoTableRowFive.getCell(1).setText(": xx");


            //两个表格之间加个换行
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun paragraphRun = paragraph.createRun();
            paragraphRun.setText("\r");


            CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
            XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);

            //添加页眉
            CTP ctpHeader = CTP.Factory.newInstance();
            CTR ctrHeader = ctpHeader.addNewR();
            CTText ctHeader = ctrHeader.addNewT();
            String headerText = "Java POI create MS word file.";
            ctHeader.setStringValue(headerText);
            XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, document);
            //设置为右对齐
            headerParagraph.setAlignment(ParagraphAlignment.RIGHT);
            XWPFParagraph[] parsHeader = new XWPFParagraph[1];
            parsHeader[0] = headerParagraph;
            policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);


            //添加页脚
            CTP ctpFooter = CTP.Factory.newInstance();
            CTR ctrFooter = ctpFooter.addNewR();
            CTText ctFooter = ctrFooter.addNewT();
            String footerText = "http://blog.csdn.net/zhouseawater";
            ctFooter.setStringValue(footerText);
            XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooter, document);
            headerParagraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFParagraph[] parsFooter = new XWPFParagraph[1];
            parsFooter[0] = footerParagraph;
            policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);


            document.write(out);
            out.close();
            System.out.println("create_table document written success.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
