package com.example.projapi.controller;

import com.example.projapi.util.ApiResult;
import com.example.projapi.util.AsposeUtil;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import java.io.*;
import java.util.*;

/**
 * @description: Aspose访问
 * @author: PCJ
 * @create: 2021-09-29 10:03
 **/
@RestController
@RequestMapping("/aspose")
public class AsposeController {

    @GetMapping("/hello")
    public ApiResult hello(){
        return ApiResult.ok("---hello---");
    }

    @PostMapping("/uploadFileAll")
    public ApiResult uploadFileAll(@RequestParam("file") MultipartFile multipartFile,HttpServletRequest request){
        long startTime = System.currentTimeMillis();
        // 截取原始文件名（包括后缀）
        String realfileName = multipartFile.getOriginalFilename();
        // 获得后缀（包括点）
        String suffix = realfileName.substring(realfileName.lastIndexOf("."));
        //Windows下文件存储地址
        String path = "C:\\Users\\Administrator\\Desktop\\compoidemo2\\";
        //Linux下文件存储地址
        //String path = "/home/images/";
        File file = new File(path);
        file.mkdirs();
        // 取一个uuid作为文件名保存（避免名字重复）写入
        String uuidfileName = UUID.randomUUID().toString() + suffix;
        System.out.println(uuidfileName);
        File f = new File(path, uuidfileName);
        BufferedOutputStream out;
        try {
            out = new BufferedOutputStream(new FileOutputStream(f));
            out.write(multipartFile.getBytes());
            out.flush();
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        long endTime = System.currentTimeMillis();
        System.out.println("上传文件耗时："+ (endTime - startTime));
        return ApiResult.ok("success!耗时："+(endTime - startTime));
    }


    @PostMapping("/wordToPdf")
//    @GetMapping("/wordToPdf")
    public String wordToPdf(String wordPath, String path, String fileName,HttpServletRequest request) throws Exception {
//        wordPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\"+wordPath;
//        pdfPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\"+pdfPath;

//        String pdfPath = "C:\\Users\\Administrator\\Desktop\\compoidemo\\文件清单2-";
        String requestPath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort();
        String pdfPath = path + fileName+".pdf";
        AsposeUtil.wordToPdf(wordPath, pdfPath);
        path = path+fileName;
        AsposeUtil.pdfToPng(pdfPath, path, requestPath);
        return "200";
    }



    @PostMapping("/upLoadReId")
    public ApiResult upLoadReId(@RequestParam("file")MultipartFile multipartFile, HttpServletRequest request) throws Exception{
        String requestPath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort();
        System.out.println(requestPath);
        // 截取原始文件名（包括后缀）
        String realfileName = multipartFile.getOriginalFilename();
        // 获得后缀（包括点）
        String suffix = realfileName.substring(realfileName.lastIndexOf("."));

        System.out.println(suffix);
        if(".doc".equals(suffix) || ".docx".equals(suffix) || ".ppt".equals(suffix) || ".pptx".equals(suffix)
        || ".xls".equals(suffix) || ".xlsx".equals(suffix)){
            //Windows下文件存储地址
            String path = "C:\\Users\\Administrator\\Desktop\\compoidemo2\\";
            //Linux下文件存储地址
//            String path = "/home/images/";
            File file = new File(path);
            file.mkdirs();
            // 取一个uuid作为文件名保存（避免名字重复）写入
            String uuidfileName = UUID.randomUUID().toString() + suffix;
            System.out.println(uuidfileName);
            File f = new File(path, uuidfileName);
            BufferedOutputStream out;
            try {
                out = new BufferedOutputStream(new FileOutputStream(f));
                out.write(multipartFile.getBytes());
                out.flush();
                out.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
            //以上为上传文件到指定目录
            String relativeFileUrl = path + uuidfileName;
            //System.out.println(path);
            System.out.println(relativeFileUrl);
            String fileName = uuidfileName.substring(0,uuidfileName.lastIndexOf("."));
            String storagePath = path+fileName;
            requestPath += "/images/"+fileName;
            //解决乱码
            //如果是windows执行，不需要加这个
            //TODO 如果是linux执行，需要添加这个*****
//            FontSettings.setFontsFolder("/usr/share/fonts",true);
            if(".doc".equals(suffix) || ".docx".equals(suffix)){
                System.out.println("------------上传的是Word文件-------------");

                String wordPath = relativeFileUrl;
                String pdfPath = path + fileName+".pdf";
                AsposeUtil.wordToPdf(wordPath, pdfPath);

                List pathList = AsposeUtil.pdfToPng(pdfPath, storagePath, requestPath);

                return ApiResult.ok(pathList);
            }else if(".ppt".equals(suffix) || ".pptx".equals(suffix)) {
                String pptPath = relativeFileUrl;
                List list = AsposeUtil.pptToImages(pptPath, storagePath, requestPath);
                return ApiResult.ok(ApiResult.ok(list));
            }else if(".xls".equals(suffix) || ".xlsx".equals(suffix)) {
                String excelPath = relativeFileUrl;
                String pdfPath = path + fileName+".pdf";
                AsposeUtil.excelTopdf(excelPath, pdfPath);

                List pathList = AsposeUtil.pdfToPng(pdfPath, storagePath, requestPath);
                return ApiResult.ok(pathList);
            }
        }

        return ApiResult.ok("--其他文件--");
    }

    @GetMapping("/wordTable")
    public ApiResult wordTable(String path){
//        String path = "C:\\Users\\Administrator\\Desktop\\compoidemo\\试验数据.docx";
        Map map = AsposeUtil.readWord(path);
        return ApiResult.ok(map);
    }


}
