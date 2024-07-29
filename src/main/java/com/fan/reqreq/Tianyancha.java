package com.fan.reqreq;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.http.HttpEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author fby
 * @apiNote
 * @date 2024/6/26
 */
public class Tianyancha {

    private static Map<String, Map<String, String>> mapMap = new HashMap<String, Map<String, String>>();
    

    public static void main(String[] args) {
        String excelFilePath = "/ISV.xlsx"; // 修改为你的Excel文件路径
        readISVNames(excelFilePath);
    }
    public static Map<String, String> findCompanyInfo(String company){
        String URL = "http://open.api.tianyancha.com/services/open/ic/baseinfo/normal?keyword=";
        String TOKEN = "";
        Map<String, String> map = new HashMap<>();
        try (CloseableHttpClient client = HttpClients.createDefault()) {
            HttpGet httpGet = new HttpGet(URL + company);
            httpGet.setHeader("Authorization", TOKEN);

            try (CloseableHttpResponse response = client.execute(httpGet)) {
                HttpEntity entity = response.getEntity();
                String result = EntityUtils.toString(entity);
                System.out.println(result);
                ObjectMapper objectMapper = new ObjectMapper();
                JsonNode jsonNode = objectMapper.readTree(result);
                JsonNode resultNode = jsonNode.path("result");
                HashMap<String, String> infoMap = new HashMap<>();
                String regCapital = resultNode.path("regCapital").asText();
                String city = resultNode.path("city").asText();
                String regLocation = resultNode.path("regLocation").asText();
                String province = regLocation.substring(0, regLocation.lastIndexOf("省")+1);
                map.put("city", city);
                map.put("province", province);
                map.put("regCapital", regCapital);
                String logFilePath = company + ".txt";
                System.out.println(result);
                writeLog(result, logFilePath);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return map;
    }

    public static List<String> readISVNames(String filePath) {
        List<String> isvNames = new ArrayList<>();
        try (InputStream is = Tianyancha.class.getResourceAsStream(filePath);
             Workbook workbook = new XSSFWorkbook(is)) {

            Sheet sheet = workbook.getSheetAt(0); // 获取第一个工作表
            Map<String, String> companyInfo = new HashMap<>();
            for (Row row : sheet) {
                if (row.getRowNum() == 0){
                    continue;
                }
                Cell cell = row.getCell(0); // 获取第一列（ISV名称列）
                if (cell != null) {
//                    isvNames.add(cell.getStringCellValue());
//                    System.out.println(cell.getStringCellValue());
                    companyInfo = findCompanyInfo(cell.getStringCellValue());
                    try {
                        Thread.sleep(200);
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }
                }
                Cell regCapitalCell = row.createCell(1); // 注册资本列
                regCapitalCell.setCellValue(companyInfo.get("regCapital"));

                Cell provinceCell = row.createCell(2); // 省份列
                provinceCell.setCellValue(companyInfo.get("province"));

                Cell cityCell = row.createCell(3); // 城市列
                cityCell.setCellValue(companyInfo.get("city"));
            }


            // 写回到文件中
            try (FileOutputStream fos = new FileOutputStream("output_ISV.xlsx")) {
                workbook.write(fos);
            }

            System.out.println("Excel文件已更新并保存为output_ISV.xlsx");

        } catch (IOException e) {
            e.printStackTrace();
        }

        return isvNames;
    }


    public static void writeISVNames(String filePath, Map<String, Map<String, String>> mapMap) {
        try (InputStream is = Tianyancha.class.getResourceAsStream(filePath);
             Workbook workbook = new XSSFWorkbook(is)) {

            Sheet sheet = workbook.getSheetAt(0); // 获取第一个工作表
            for (Row row : sheet) {
                if (row.getRowNum() == 0){
                    continue;
                }
                Cell capitalCell = row.createCell(1); // 注册资本列
                capitalCell.setCellValue("200万");

                Cell provinceCell = row.createCell(2); // 省份列
                provinceCell.setCellValue("山东");

                Cell cityCell = row.createCell(3); // 城市列
                cityCell.setCellValue("济南");
            }

            // 写回到文件中
            try (FileOutputStream fos = new FileOutputStream("1output_ISV.xlsx")) {
                workbook.write(fos);
            }

            System.out.println("Excel文件已更新并保存为output_ISV.xlsx");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static void writeLog(String message, String filePath) {
        String path = "d:/company/" + filePath;
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(path, true))) {
            writer.write(message);
            writer.newLine();
            System.out.println("日志已成功写入文件");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
