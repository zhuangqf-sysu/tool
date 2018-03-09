package cn.zhuangqf.tool.excel;

import com.asiainfo.ngbomc.common.WebResult;
import com.google.gson.Gson;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLConnection;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author zhuangqf
 * @date 2018/2/26
 */
@Controller
@RequestMapping("/common/excel")
public class ExcelController {
    private static final Logger logger = LoggerFactory.getLogger(ExcelController.class);

    @RequestMapping("export")
    @ResponseBody
    public void export(@RequestBody ExportRequest requestBody, HttpServletResponse response) throws Exception {
        if (logger.isDebugEnabled()) {
            logger.debug(new Gson().toJson(requestBody));
        }

        response = setResponseInfo(response, requestBody.getFileName());

        ExcelService.createExportWorkbook(requestBody.getFileName())
                .addSheet(requestBody.getSheetName())
                .addHeaders(requestBody.getHeaders())
                .addData(requestBody.getData())
                .save(response.getOutputStream());

    }

    @RequestMapping("exportMultiple")
    @ResponseBody
    public void export(@RequestBody List<ExportRequest> requestBody, HttpServletResponse response) throws Exception {
        if (logger.isDebugEnabled()) {
            logger.debug(new Gson().toJson(requestBody));
        }

        if (requestBody.isEmpty()) {
            throw new Exception("无导出表格");
        }

        String fileName = requestBody.get(0).getFileName();
        response = setResponseInfo(response, fileName);

        ExportWorkbook workbook = ExcelService.createExportWorkbook(fileName);

        requestBody.forEach(sheet -> workbook.addSheet(sheet.getSheetName())
                .addHeaders(sheet.getHeaders())
                .addData(sheet.getData()));
        workbook.save(response.getOutputStream());
    }

    private HttpServletResponse setResponseInfo(HttpServletResponse response, String fileName) throws UnsupportedEncodingException {
        //判断文件类型
        String mimeType = URLConnection.guessContentTypeFromName(fileName);
        if (mimeType == null) {
            mimeType = "application/octet-stream";
        }
        response.setContentType(mimeType);

        //文件名编码，解决乱码问题
        String encodedFileName = URLEncoder.encode(fileName, "utf-8").replaceAll("\\+", "%20");
//        String userAgentString = request.getHeader("User-Agent");
//        String browser = UserAgent.parseUserAgentString(userAgentString).getBrowser().getGroup().getName();
//        if (browser.equals("Chrome") || browser.equals("Internet Exploer") || browser.equals("Safari")) {
//            encodedFileName = URLEncoder.encode(fileName, "utf-8").replaceAll("\\+", "%20");
//        } else {
//            encodedFileName = MimeUtility.encodeWord(fileName);
//        }

        //设置Content-Disposition响应头，一方面可以指定下载的文件名，另一方面可以引导浏览器弹出文件下载窗口
        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"; filename*=utf-8''" + fileName);

        //设置Content-Disposition响应头，一方面可以指定下载的文件名，另一方面可以引导浏览器弹出文件下载窗口
        response.setHeader("Content-Disposition", "attachment;fileName=\"" + encodedFileName + "\"");

        return response;
    }

    @RequestMapping("create")
    @ResponseBody
    public WebResult create(@RequestBody ExportRequest requestBody) {
        if (logger.isDebugEnabled()) {
            logger.debug(new Gson().toJson(requestBody));
        }
        try {
            File file = File.createTempFile("EXPORT_", requestBody.getFileName());
            logger.debug(file.getAbsolutePath());
            try (OutputStream outputStream = new FileOutputStream(file)) {
                ExcelService.createExportWorkbook(requestBody.getFileName()).addSheet()
                        .addHeaders(requestBody.getHeaders())
                        .addData(requestBody.getData())
                        .save(outputStream);
                outputStream.flush();
                Map<String, String> data = new HashMap<>(1);
                data.put("fileId", file.getName().replace(requestBody.getFileName(), ""));
                data.put("fileName", requestBody.getFileName());
                return new WebResult().success("", data);
            }
        } catch (Exception e) {
            logger.error(e.getMessage());
            return WebResult.failure(e.getMessage());
        }
    }

    @RequestMapping("download")
    @ResponseBody
    public void download(String fileName, String fileId, HttpServletResponse response, HttpServletRequest request) throws IOException {
        logger.debug(fileId);
        logger.debug(fileName);
        String path = File.createTempFile("test", "").getParent();
        File file = new File(path, fileId + fileName);

        //判断文件是否存在
        if (!file.exists()) {
            return;
        }

        setResponseInfo(response, fileName);

        //文件下载
        InputStream in = new BufferedInputStream(new FileInputStream(file));
        FileCopyUtils.copy(in, response.getOutputStream());
    }
}
