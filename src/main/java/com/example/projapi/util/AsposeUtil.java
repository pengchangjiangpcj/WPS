package com.example.projapi.util;

import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.slides.ISlide;
import com.aspose.slides.ISlideCollection;
import com.aspose.slides.Presentation;
import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.icepdf.core.util.GraphicsRenderingHints;

import javax.imageio.ImageIO;

/**
 * @description: Aspose工具类
 * @author: PCJ
 * @create: 2021-09-29 09:50
 **/
public class AsposeUtil implements Serializable{

    /**
     * 加载license 用于破解 不生成水印
     */
    private static void getLicense() {
        try (InputStream is = AsposeUtil.class.getClassLoader().getResourceAsStream("License.xml")) {
            License license = new License();
            license.setLicense(is);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     *word转pdf
     * @param wordPath word文件保存的路径
     * @param pdfPath 转换后pdf文件保存的路径
     */
    public static void wordToPdf(String wordPath, String pdfPath) throws Exception{
        //wordPath word文件保存的路径
//        String wordPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\文件清单.docx";
        //pdfPath  转换后pdf文件保存的路径
//        String pdfPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\文件清单2.pdf";
        getLicense();
        File file = new File(pdfPath);

        long old = System.currentTimeMillis();
        try (FileOutputStream os = new FileOutputStream(file)) {
            Document doc = new Document(wordPath);
            doc.save(os, SaveFormat.PDF);
            long now = System.currentTimeMillis();
            System.out.println("pdf转换成功，共耗时：" + ((now - old) / 1000.0) + "秒"); // 转化用时
        }
    }

    public static List pdfToPng(String pdfPath, String path, String requestPath) throws Exception{

//        pdfPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\文件清单2.pdf";
//        path = "C:\\Users\\Administrator\\Desktop\\compoidemo\\文件清单2-";
        List pathList = new ArrayList();
        org.icepdf.core.pobjects.Document document = new org.icepdf.core.pobjects.Document();
        document.setFile(pdfPath);
        //缩放比例
        float scale = 2.5f;
        //旋转角度
        float rotation = 0f;

        for (int i = 0; i < document.getNumberOfPages(); i++) {
            BufferedImage image = (BufferedImage)
                    document.getPageImage(i, GraphicsRenderingHints.SCREEN, org.icepdf.core.pobjects.Page.BOUNDARY_CROPBOX, rotation, scale);
            RenderedImage rendImage = image;
            try {
                String imgName = i +1+ ".png";

                File file = new File(path + imgName);
                System.out.println(requestPath + imgName);
                pathList.add(requestPath + imgName);
                ImageIO.write(rendImage, "png", file);
            } catch (IOException e) {
                e.printStackTrace();
            }
            image.flush();
        }
        document.dispose();

        return pathList;
    }


    /**
     * @param sourceFilePath 资源文件路径
     * @param storagePath 图像存放地址
     * @param requestPath 图像存放地址
     * @return
     */
    public static List pptToImages(String sourceFilePath, String storagePath, String requestPath) {
        List pathList = new ArrayList();
        long startTime = System.currentTimeMillis();

        getLicensePpt();
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
                System.out.println(requestPath + imgName);
                pathList.add(requestPath + imgName);
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
        return pathList;
    }
    private static void getLicensePpt() {
        try (InputStream is = AsposeUtil.class.getClassLoader().getResourceAsStream("License.xml")) {
            com.aspose.slides.License license = new com.aspose.slides.License();
            license.setLicense(is);
        } catch (IOException e) {
            e.printStackTrace();
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



    public static Map readWord(String path) {
        String buffer = "";
        List<String> list = new ArrayList();
        try {
            if (path.endsWith(".doc")) {
                InputStream is = new FileInputStream(new File(path));
                WordExtractor ex = new WordExtractor(is);
                buffer = ex.getText();
                ex.close();
            } else if (path.endsWith("docx")) {
                OPCPackage opcPackage = POIXMLDocument.openPackage(path);
//                POIXMLTextExtractor extractor = new XWPFWordExtractor(opcPackage);
                XWPFWordExtractor extractor1 = new XWPFWordExtractor(opcPackage);

                XWPFDocument document = (XWPFDocument) extractor1.getDocument();
                List<XWPFTable> tables = document.getTables();
                for (int i = 0; i < tables.size(); i++) {
                    System.out.println(tables.get(i).getText());
                    System.out.println("-------------------------------------------------");
                    list.add(tables.get(i).getText());
                }
                buffer = extractor1.getText();
                extractor1.close();
            } else {
                System.out.println("此文件不是word文件！");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println(buffer);
        Map map = new HashMap<>(2);
        map.put("buffer",buffer);
        map.put("tables",list);
        return map;
    }

}
