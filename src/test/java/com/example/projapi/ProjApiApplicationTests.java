package com.example.projapi;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.slides.ISlide;
import com.aspose.slides.ISlideCollection;
import com.aspose.slides.Presentation;
import com.aspose.words.Document;
import com.aspose.words.Range;
import com.example.projapi.util.AsposeUtil;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.TextRun;
import org.apache.poi.hslf.usermodel.*;
import org.apache.poi.hwpf.extractor.WordExtractor;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;

import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.springframework.boot.test.context.SpringBootTest;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;


@SpringBootTest
class ProjApiApplicationTests {

    @Test
    void contextLoads() {

        try{
            //没有参数
            File fileCreateByNo=new File("");
            System.out.println("fileCreateByNo=="+fileCreateByNo);
            System.out.println("fileCreateByNo=="+fileCreateByNo.getCanonicalPath());
            System.out.println("-----------------------------");
            //一个点的参数
            File fileCreateByPoint=new File(".");
            System.out.println("fileCreateByPoint=="+fileCreateByPoint);
            System.out.println("fileCreateByPoint=="+fileCreateByPoint.getCanonicalPath());
            System.out.println("-----------------------------");
            //两个点的参数
            File fileTwoPoint = new File("..");
            System.out.println("fileTwoPoint=="+fileTwoPoint);
            System.out.println("fileTwoPoint=="+fileTwoPoint.getCanonicalPath());
            System.out.println("-----------------------------");
            //一个点两条斜线的参数
            File filePLL = new File(".\\");
            System.out.println("filePLL=="+filePLL);
            System.out.println("filePLL=="+filePLL.getCanonicalPath());
            System.out.println("-----------------------------");
            //当前工作目录
            String currentWorkPath=System.getProperty("user.dir");
            System.out.println("currentWorkPath=="+currentWorkPath);
        }catch(Exception e){
            System.out.println("IOException....出问题咯");
        }
    }

    /**
     * 删除文件
     */
    @Test
    void contextLoads2() {
        System.out.println(new Date());
        //要删除的文件夹或文件
//        delAllFile("C:\\Users\\Administrator\\Desktop\\compoidemo2");
        delFolder("C:\\Users\\Administrator\\Desktop\\compoidemo2\\文件清单1.png");
        System.out.println(new Date());
    }
    public boolean delAllFile(String path) {
        boolean flag = false;
        File file = new File(path);
        if (!file.exists()) {
            return flag;
        }
        if (!file.isDirectory()) {
            return flag;
        }
        String[] tempList = file.list();
        File temp = null;
        for (int i = 0; i < tempList.length; i++) {
            if (path.endsWith(File.separator)) {
                temp = new File(path + tempList[i]);
            } else {
                temp = new File(path + File.separator + tempList[i]);
            }
            if (temp.isFile()) {
                temp.delete();
            }
            if (temp.isDirectory()) {
                delAllFile(path + "/" + tempList[i]);//先删除文件夹里面的文件
                //delFolder(path + "/" + tempList[i]);//再删除空文件夹
                flag = true;
            }
        }
        return flag;
    }
    public void delFolder(String folderPath) {
        try {
            delAllFile(folderPath); //删除完里面所有内容
            String filePath = folderPath;
            filePath = filePath.toString();
            java.io.File myFilePath = new java.io.File(filePath);
            myFilePath.delete(); //删除空文件夹
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     *ppt转图像 poi-3.8版本
     */
    @Test
    void contextLoads3() {
        // 读入PPT文件
//        File file = new File("C:\\Users\\Administrator\\Desktop\\compoidemo\\泵间基础知识.pptx");
//        File file = new File("C:\\Users\\Administrator\\Desktop\\compoidemo\\操作技能（常规加注泵间）.ppt");
        String pptPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\操作技能（常规加注泵间）.ppt";
//        String pptxPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\泵间基础知识.pptx";
        String storagePath = "C:\\Users\\Administrator\\Desktop\\compoidemo2\\";
        String uuid = UUID.randomUUID().toString();
        doPPTtoImage(pptPath,storagePath+uuid);

    }

    public static boolean doPPTtoImage(String pptPath,String storagePath) {
        File file = new File(pptPath);
        boolean isppt = checkFile(file);
        if (!isppt) {
            System.out.println("The image you specify don't exit!");
            return false;
        }
        try {
            FileInputStream is = new FileInputStream(file);
            SlideShow ppt = new SlideShow(is);
            //及时关闭掉 输入流
            is.close();

            Dimension pgsize = ppt.getPageSize();
            Slide[] slide = ppt.getSlides();
            for (int i = 0; i < slide.length; i++) {
                TextRun[] truns = slide[i].getTextRuns();
                for (int k = 0; k < truns.length; k++) {
                    RichTextRun[] rtruns = truns[k].getRichTextRuns();
                    for (int l = 0; l < rtruns.length; l++) {
                        // 重新设置 字体索引 和 字体名称 是为了防止生成的图片乱码问题
                        rtruns[l].setFontIndex(1);
                        rtruns[l].setFontName("宋体");
                    }
                }
                //根据幻灯片大小生成图片
                BufferedImage img = new BufferedImage(pgsize.width, pgsize.height,
                        BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();

                graphics.setPaint(Color.BLUE);
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
                slide[i].draw(graphics);
                // 这里设置图片的存放路径和图片的格式(jpeg,png,bmp等等),注意生成文件路径
//                File path = new File("C:\\Users\\Administrator\\Desktop\\compoidemo2");
//                if (!path.exists()) {
//                    path.mkdir();
//                }
                String imgName = i +1+ ".png";
                File path = new File(storagePath+imgName);
                System.out.println(storagePath+imgName);
                // 可测试多种图片格式
                FileOutputStream out = new FileOutputStream(path);
                javax.imageio.ImageIO.write(img, "png", out);
                out.close();
            }
            System.out.println("success!!");
            return true;
        } catch (FileNotFoundException e) {
            System.out.println(e);
        } catch (IOException e) {

        }
        return false;
    }

    // function 检查文件是否为PPT
    public static boolean checkFile(File file) {
        boolean isppt = false;
        String filename = file.getName();
        String suffixname = null;
        if (filename != null && filename.indexOf(".") != -1) {
            suffixname = filename.substring(filename.lastIndexOf("."));
            if (suffixname.equals(".ppt") || suffixname.equals(".pptx")) {
                isppt = true;
            }
            return isppt;
        } else {
            return isppt;
        }
    }
    @Test
    void contextLoads4() {
        String pptxPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\泵间基础知识.pptx";
        String storagePath = "C:\\Users\\Administrator\\Desktop\\compoidemo2\\";
        String uuid = UUID.randomUUID().toString();
        doPPT2007toImage(pptxPath, storagePath+uuid);
    }

    /**
     * ppt2007文档的转换 后缀为.pptx
     * @param pptPath PPTX文件地址
     * @param storagePath 图片将要保存的路径目录（不是文件）
     * @return
     */
    public static boolean doPPT2007toImage(String pptPath,String storagePath) {
        File pptFile = new File(pptPath);

        FileInputStream is = null ;


        try {

            is = new FileInputStream(pptFile);

            XMLSlideShow xmlSlideShow = new XMLSlideShow(is);

            is.close();

            // 获取大小
            Dimension pgsize = xmlSlideShow.getPageSize();

            // 获取幻灯片
            XSLFSlide[] slides = xmlSlideShow.getSlides();

            for (int i = 0 ; i < slides.length ; i++) {

                XSLFShape[] shapes = slides[i].getShapes();
                for (XSLFShape shape : shapes) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape xslfTextShape = (XSLFTextShape) shape;
                        for (XSLFTextParagraph xslfTextParagraph : xslfTextShape) {
                            for (XSLFTextRun xslfTextRun : xslfTextParagraph) {
                                xslfTextRun.setFontFamily("宋体");
                            }
                        }
                    }
                }

                //根据幻灯片大小生成图片
                BufferedImage img = new BufferedImage(pgsize.width,pgsize.height, BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();

                graphics.setPaint(Color.white);
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width,pgsize.height));

                // 最核心的代码
                slides[i].draw(graphics);

                //图片将要存放的路径
                String imgName = i +1+ ".png";
                String absolutePath = storagePath+imgName;
                File jpegFile = new File(absolutePath);
                //如果图片存在，则不再生成
//                if (jpegFile.exists()) {
//                    continue;
//                }
                // 这里设置图片的存放路径和图片的格式(jpeg,png,bmp等等),注意生成文件路径
                FileOutputStream out = new FileOutputStream(jpegFile);

                // 写入到图片中去
                ImageIO.write(img, "png", out);
                System.out.println(absolutePath);
                out.close();

            }
            System.out.println("PPTX转换成图片 成功！");
            return true;
        } catch (Exception e) {
            System.out.println("PPTX转换成图片 发生异常！");
        }

        return false;
    }
    @Test
    void contextLoads5() {
        String pptxPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\泵间基础知识.pptx";
        String pptPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\操作技能（常规加注泵间）.ppt";
        String storagePath = "C:\\Users\\Administrator\\Desktop\\compoidemo2\\";
        String uuid = UUID.randomUUID().toString();
        convertPptToImages(pptxPath, storagePath+uuid);
    }

    /**
     * @param sourceFilePath 资源文件路径
     * @param storagePath 图像存放地址
     * @return
     */
    public static void convertPptToImages(String sourceFilePath, String storagePath) {

        long startTime = System.currentTimeMillis();

        getLicense();
        Presentation pres = new Presentation(sourceFilePath);
        ISlideCollection slides = pres.getSlides();
        int idx = 1;
        for (int i = 0; i < slides.size(); i++) {
            ISlide slide = slides.get_Item(i);
            int height = (int) (pres.getSlideSize().getSize().getHeight());
            int width = (int) (pres.getSlideSize().getSize().getWidth());
            BufferedImage img = slide.getThumbnail(new java.awt.Dimension(width, height));
            FileOutputStream out = null;
            File pngFile = null;
            try {
                String imgName = i +1+ ".png";
                pngFile = new File(storagePath+imgName);
                out = new FileOutputStream(pngFile);
                ImageIO.write(img, "png", out);
                System.out.println(storagePath+imgName);
            } catch (Exception e) {

            } finally {
                try {
                    if (out != null) {
                        out.flush();
                        out.close();
                    }
                    if (img != null) {
                        img.flush();
                    }
                } catch (IOException e) {

                }
            }
        }
        long endTime = System.currentTimeMillis();
        System.out.println("PPT转PNG耗时："+ (endTime - startTime));
    }
    private static void getLicense() {
        try (InputStream is = AsposeUtil.class.getClassLoader().getResourceAsStream("License.xml")) {
            com.aspose.slides.License license = new com.aspose.slides.License();
            license.setLicense(is);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 测试Excel转pdf
     */
    @Test
    void contextLoads6() {
        String xlsPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\养老金代发机构与社保卡银行.xls";
        String xlsxPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\示例-家庭收入支出表.xlsx";
        String storagePath = "C:\\Users\\Administrator\\Desktop\\compoidemo2\\";
        String uuid = UUID.randomUUID().toString();
        String pdfPath = storagePath + uuid+".pdf";
//        excel2pdf(xlsxPath, pdfPath);
        excelTopdf(xlsPath, pdfPath);
    }

    /**
     * @param excelPath 需要被转换的excel全路径带文件名
     * @param pdfPath 转换之后pdf的全路径带文件名
     */
    public static void excel2pdf(String excelPath, String pdfPath) {
        getLicenseExcel();
        try {
            long old = System.currentTimeMillis();
            Workbook wb = new Workbook(excelPath);// 原始excel路径
            FileOutputStream fileOS = new FileOutputStream(new File(pdfPath));
            wb.save(fileOS, com.aspose.cells.SaveFormat.PDF);
            fileOS.close();
            long now = System.currentTimeMillis();
            System.out.println("共耗时：" + ((now - old) / 1000.0) + "秒"); //转化用时
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void getLicenseExcel() {
        try (InputStream is = AsposeUtil.class.getClassLoader().getResourceAsStream("License.xml")) {
            com.aspose.cells.License license = new com.aspose.cells.License();
            license.setLicense(is);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * excel 转为pdf 输出。
     *
     * @param sourceFilePath  excel文件
     * @param desFilePathd  pad 输出文件目录
     */
    public static void excelTopdf(String sourceFilePath, String desFilePathd ){
        getLicenseExcel();
        try {
            Workbook wb = new Workbook(sourceFilePath);// 原始excel路径

            FileOutputStream fileOS = new FileOutputStream(desFilePathd);
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.setOnePagePerSheet(true);


            int[] autoDrawSheets={3};
            //当excel中对应的sheet页宽度太大时，在PDF中会拆断并分页。此处等比缩放。
            autoDraw(wb,autoDrawSheets);

            int[] showSheets={0};
            //隐藏workbook中不需要的sheet页。
//            printSheetPage(wb,showSheets);
            wb.save(fileOS, pdfSaveOptions);
            fileOS.flush();
            fileOS.close();
            System.out.println("完毕");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /**
     * 设置打印的sheet 自动拉伸比例
     * @param wb
     * @param page 自动拉伸的页的sheet数组
     */
    public static void autoDraw(Workbook wb,int[] page){
        if(null!=page&&page.length>0){
            for (int i = 0; i < page.length; i++) {
                wb.getWorksheets().get(i).getHorizontalPageBreaks().clear();
                wb.getWorksheets().get(i).getVerticalPageBreaks().clear();
            }
        }
    }


    /**
     * 隐藏workbook中不需要的sheet页。
     * @param wb
     * @param page 显示页的sheet数组
     */
    public static void printSheetPage(Workbook wb,int[] page){
        for (int i= 1; i < wb.getWorksheets().getCount(); i++)  {
            wb.getWorksheets().get(i).setVisible(false);
        }
        if(null==page||page.length==0){
            wb.getWorksheets().get(0).setVisible(true);
        }else{
            for (int i = 0; i < page.length; i++) {
                wb.getWorksheets().get(i).setVisible(true);
            }
        }
    }


    @Test
    void contextLoads7() throws Exception {
        //wordPath word文件保存的路径
        String wordPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\试验数据.docx";
        //pdfPath  转换后pdf文件保存的路径
        String pdfPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\试验数据.pdf";

        wordToPdf(wordPath,pdfPath);
    }
    public static void wordToPdf(String wordPath, String pdfPath) throws Exception{

        getLicense();
        File file = new File(pdfPath);

        long old = System.currentTimeMillis();
        try (FileOutputStream os = new FileOutputStream(file)) {
            Document doc = new Document(wordPath);
            Range range = doc.getRange();
            System.out.println(range);
//            doc.save(os, SaveFormat.PDF);
            long now = System.currentTimeMillis();
            System.out.println("pdf转换成功，共耗时：" + ((now - old) / 1000.0) + "秒"); // 转化用时
        }
    }

    @Test
    void contextLoads8(){
        //wordPath word文件保存的路径
        String wordPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\试验数据.docx";
        String readWord = readWord(wordPath);

        System.out.println(readWord);
    }

    public static String readWord(String path) {
        String buffer = "";
        try {
            if (path.endsWith(".doc")) {
                InputStream is = new FileInputStream(new File(path));
                WordExtractor ex = new WordExtractor(is);
                buffer = ex.getText();
                ex.close();
            } else if (path.endsWith("docx")) {
                OPCPackage opcPackage = POIXMLDocument.openPackage(path);
                POIXMLTextExtractor extractor = new XWPFWordExtractor(opcPackage);
                XWPFWordExtractor extractor1 = new XWPFWordExtractor(opcPackage);

                XWPFDocument document = (XWPFDocument) extractor1.getDocument();
                List<XWPFTable> tables = document.getTables();
                System.out.println(tables.size());
                buffer = extractor1.getText();
                extractor1.close();
            } else {
                System.out.println("此文件不是word文件！");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        return buffer;
    }


    @Test
    void contextLoads9() throws Exception{
        String wordPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\试验数据.docx";
        FileInputStream is = new FileInputStream(wordPath);
        XWPFDocument document = new XWPFDocument(is);
        try {
            // 获取word中的所有段落与表格
            List<IBodyElement> elements = document.getBodyElements();
            for (IBodyElement element : elements) {
                // 段落
                if (element instanceof XWPFParagraph) {
                    getParagraphText((XWPFParagraph) element);
                }
                // 表格
                else if (element instanceof XWPFTable) {
                    getTabelText((XWPFTable) element);
                }
            }
        } finally {
            is.close();
        }

    }


    /**
     * 获取段落内容
     *
     * @param paragraph
     */
    private void getParagraphText(XWPFParagraph paragraph) {
        // 获取段落中所有内容
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs.size() == 0) {
//            System.out.println("按了回车（新段落）");
            return;
        }
        StringBuffer runText = new StringBuffer();
        for (XWPFRun run : runs) {
            runText.append(run.text());
        }
        if (runText.length() > 0) {
//            runText.append("，对齐方式：").append(paragraph.getAlignment().name());
            System.out.println(runText);
        }
    }

    /**
     * 获取表格内容
     *
     * @param table
     */
    private void getTabelText(XWPFTable table) {
        List<XWPFTableRow> rows = table.getRows();

        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();

            for (XWPFTableCell cell : cells) {
                // 简单获取内容（简单方式是不能获取字体对齐方式的）
                // System.out.println(cell.getText());
//                CTTcPr tcPr = cell.getCTTc().getTcPr();
                // 一个单元格可以理解为一个word文档，单元格里也可以加段落与表格
                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                for (XWPFParagraph paragraph : paragraphs) {
                    getParagraphText(paragraph);
                }
            }
        }
    }


}
