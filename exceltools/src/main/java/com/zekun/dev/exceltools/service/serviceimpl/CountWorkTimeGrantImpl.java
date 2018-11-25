package com.zekun.dev.exceltools.service.serviceimpl;

import com.zekun.dev.exceltools.model.ExceltFileModel;
import com.zekun.dev.exceltools.service.CountWorkTimeGrant;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class CountWorkTimeGrantImpl implements CountWorkTimeGrant {
    private Logger logger = LoggerFactory.getLogger(CountWorkTimeGrantImpl.class);
    private int index_name = 0;
    private int index_workNum = 1;
    private int index_date = 2;
    private int index_startTime = 3;
    private int index_endTime = 4;
    @Override
    public String countResultFromExcel(ExceltFileModel exceltFileModel) {
        int grant = exceltFileModel.getGrant();
        String countTimePoint = exceltFileModel.getCountTimePoint();
        MultipartFile namesFile = exceltFileModel.getNamesFile();
        MultipartFile workHours = exceltFileModel.getWorkHours();
        String savePath = "D:\\晚餐补助统计结果_zekun";
        String resultPath = "";
        try {
            if (0 >= namesFile.getSize()) {
                resultPath = "请先上传统计人员名单";
            }
            if (0 >= workHours.getSize()) {
                resultPath = "请先上传本月总工时表";
            }

            int hoursPoint = Integer.valueOf(countTimePoint.split(":")[0]);
            int minutePoint = Integer.valueOf(countTimePoint.split(":")[1]);

            TreeMap<String, String> namesWorkNumMap = getNameWorkNumMap(namesFile);

            HashMap<String, Integer> countResult = new HashMap<String, Integer>();
            ArrayList<String[]> workTimeBasisList = new ArrayList<String[]>();

            countExcel(workHours, hoursPoint, minutePoint, namesWorkNumMap, countResult, workTimeBasisList);

            File tempFile = new File(savePath);
            if (!tempFile.exists()) {
                tempFile.mkdirs();
            }

            String fileName = savePath + "\\晚餐补助统计结果表.xls";
            writeCountResultXls(fileName, grant, namesWorkNumMap, countResult);

            fileName = savePath + "\\晚餐补助工时统计依据表.xls";
            writeTimeBasisXls(fileName, workTimeBasisList);

            resultPath = savePath;
        } catch (IOException e) {
            resultPath = "程序出现异常^-^只支持处理03版Excel";
            e.printStackTrace();
        }

        return resultPath;
    }

    private void countExcel(MultipartFile workHours, int hoursPoint, int minutePoint, TreeMap<String, String> namesWorkNumMap, HashMap<String, Integer> countResult, ArrayList<String[]> workTimeBasisList) throws IOException {
        HSSFWorkbook book = new HSSFWorkbook(workHours.getInputStream());
        HSSFSheet sheet = book.getSheetAt(1);
        int rowNum = sheet.getLastRowNum();

        for (int i = 1; i <= rowNum; i++) {
            HSSFRow row = sheet.getRow(i);// 行数
            HSSFCell workNumCell = row.getCell(index_workNum);
            workNumCell.setCellType(Cell.CELL_TYPE_STRING);
            String workNum = getValue(workNumCell);
            if (namesWorkNumMap.containsKey(workNum)) {
                HSSFCell endTimeCell = row.getCell(index_endTime);
                String endTime = getValue(endTimeCell);
                if (endTime.contains(":")) {
                    String[] times = endTime.split(":");
                    int hours = Integer.valueOf(times[0]);
                    int minute = Integer.valueOf(times[1]);
                    if ((hours > hoursPoint || 2 > hours) || (hours == hoursPoint &&  minute >= minutePoint)) {
                        if (!countResult.containsKey(workNum)) {
                            countResult.put(workNum, 0);
                        }
                        countResult.put(workNum, (countResult.get(workNum) + 1));
                    }
                }

                String name = row.getCell(index_name).getStringCellValue();
                row.getCell(index_date).setCellType(Cell.CELL_TYPE_STRING);
                String date = row.getCell(index_date).getStringCellValue();
                HSSFCell startTimeCell = row.getCell(index_startTime);
                String startTime = getValue(startTimeCell);
                String[] workTimeBasis = new String[5];
                workTimeBasis[0] = name;
                workTimeBasis[1] = workNum;
                workTimeBasis[2] = date;
                workTimeBasis[3] = startTime;
                workTimeBasis[4] = endTime;
                workTimeBasisList.add(workTimeBasis);
            }
        }
    }

    private TreeMap<String, String> getNameWorkNumMap(MultipartFile namesFile) throws IOException {
        TreeMap<String, String> namesWorkNumMap = new TreeMap<String, String>();
        HSSFWorkbook book = new HSSFWorkbook(namesFile.getInputStream());
        HSSFSheet sheet = book.getSheetAt(0);
        int rowNum = sheet.getLastRowNum();
        for (int i = 1; i <= rowNum; i++) {
            HSSFRow row = sheet.getRow(i);// 行数
            if(row.getCell(index_name) == null){//姓名为空的行跳过
                continue;
            }else if(row.getCell(index_workNum) == null){//工号为空的行跳过
                continue;
            }else{
                HSSFCell nameCell = row.getCell(index_name);
                nameCell.setCellType(Cell.CELL_TYPE_STRING);
                String name = getValue(nameCell);

                HSSFCell workNumCell = row.getCell(index_workNum);
                workNumCell.setCellType(Cell.CELL_TYPE_STRING);
                String workNum = getValue(workNumCell);
                namesWorkNumMap.put(workNum, name);
            }
        }
        return namesWorkNumMap;

    }

    private void writeTimeBasisXls(String fileName, List<String[]> list) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("0");
        HSSFRow rowNumOne = sheet.createRow(0);
        rowNumOne.createCell(0).setCellValue("姓名");
        rowNumOne.createCell(1).setCellValue("工号");
        rowNumOne.createCell(2).setCellValue("日期");
        rowNumOne.createCell(3).setCellValue("上班时间");
        rowNumOne.createCell(4).setCellValue("下班时间");
        for(int i = 0; i < list.size(); i++) {
            HSSFRow row = sheet.createRow(i + 1);
            String[] str = list.get(i);
            int size = str.length;
            for (int j = 0; j < size; j++) {
                HSSFCell cell = row.createCell(j);
                if (1 == j) {
                    cell.setCellValue(Integer.valueOf(str[j]));
                    continue;
                }
                cell.setCellValue(str[j]);
            }
        }
        FileOutputStream outputStream = new FileOutputStream(fileName);
        wb.write(outputStream);
        outputStream.close();
    }

    private void writeCountResultXls(String fileName, int grant, TreeMap<String, String> namesWorkNumMap, HashMap<String, Integer> countResult) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("0");
        HSSFRow rowNumOne = sheet.createRow(0);
        rowNumOne.createCell(0).setCellValue("姓名");
        rowNumOne.createCell(1).setCellValue("工号");
        rowNumOne.createCell(2).setCellValue("晚餐补助次数");
        rowNumOne.createCell(3).setCellValue("晚餐补助金额(元)");

        Iterator<Map.Entry<String, String>> entries = namesWorkNumMap.entrySet().iterator();
        int i = 1;
        while (entries.hasNext()) {
            Map.Entry<String, String> entry = entries.next();
            String key = entry.getKey();
            if (countResult.containsKey(key)) {
                HSSFRow row = sheet.createRow(i);
                HSSFCell cell0 = row.createCell(0);
                cell0.setCellValue(entry.getValue());
                HSSFCell cell1 = row.createCell(1);
                cell1.setCellValue(Integer.valueOf(key));
                HSSFCell cell2 = row.createCell(2);
                cell2.setCellValue(countResult.get(key));
                HSSFCell cell3 = row.createCell(3);
                cell3.setCellValue(countResult.get(key) * grant);
                i++;
            }
        }
        FileOutputStream outputStream = new FileOutputStream(fileName);
        wb.write(outputStream);
        outputStream.close();
    }

    private String getValue(HSSFCell hssfCell) {
        if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
            return String.valueOf(hssfCell.getBooleanCellValue());
        } else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(hssfCell)) {// 判断是否为日期类型
                Date date = hssfCell.getDateCellValue();
                DateFormat formater = new SimpleDateFormat(
                        "HH:mm");
                return formater.format(date);
            }
            return String.valueOf(hssfCell.getNumericCellValue());
        } else {
            return String.valueOf(hssfCell.getStringCellValue());
        }
    }

}
