package org.lszjaf.utils.excel;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.lszjaf.utils.common.StringUtil;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;
/**
 * 目前该工具类支持：
 * 1.Excel的读取，仅限 Excel的内容可以映射为一个Java对象的情况。
 * 2.Excel的写入，仅限 可以把一个Java对象写入一个Excel文件的情况。
 * 备注：（Excel的列名 要和Java对象的属性名称一一对应）
 * @author Joybana
 * @date 2019-01-17 10:15:04
 */
public class ExcelUtil {

    private static final String GET = "get";
    private static final String SET = "set";



    /**
     * read excel file transfer it to a java object in the specified sheet
     * 推荐使用
     *
     * @param filepath
     * @param receiveType 读取的Excel要转变成那个Java对象
     * @param sheetNum    from zero start, zero is the first sheet!
     * @param <K>
     * @return
     * @throws Exception
     */
    public static <K> List<K> read(String filepath, Class<K> receiveType, int sheetNum) throws Exception {
        if (filepath == null || filepath.trim().isEmpty() || receiveType == null || receiveType == Void.class) {
            return null;
        }
        Workbook wb = readExcel(filepath);
        if (wb == null) {
            return null;
        }
        //获取excel表有几张工作簿
        int sheetSize = wb.getNumberOfSheets();
        if (sheetNum > sheetSize || sheetNum < 0) {
            return null;
        }
        //用来存放表中数据
        List list = new ArrayList<>();

        //获取sheet
        Sheet sheet = wb.getSheetAt(sheetNum);
        //获取最大行数
        int rowMax = sheet.getPhysicalNumberOfRows();
        if (rowMax < 1) {
            return null;
        }
        //获取第一行
        Row rowFirst = sheet.getRow(0);
        //获取最大列数
        int columnMax = rowFirst.getPhysicalNumberOfCells();
        if (columnMax < 1) {
            return null;
        }

        Row row = null;
        for (int k = 1; k < rowMax; k++) {
            row = sheet.getRow(k);
            if (row == null) {
                continue;
            }

            //每一行都是一个Java对象
            Object object = receiveType.newInstance();

            for (int j = 0; j < columnMax; j++) {
                //获取字段名称
                String fieldName = rowFirst.getCell(j).getStringCellValue();
//                    System.out.println("获取到的字段名=" + fieldName);
                fieldName = StringUtil.toUpperFirstLetter(fieldName);
                //通get方法获取字段的类型
                Method methodGet = receiveType.getDeclaredMethod(GET + fieldName, null);
                Class type = methodGet.getReturnType();
                //调用对应的set方法注入数值
                Method methodSet = receiveType.getDeclaredMethod(SET + fieldName, type);
                methodSet.invoke(object, getCellFormatValue(row.getCell(j), type));
            }
            list.add(object);
        }
        return list;
    }


    /**
     * read excel file transfer it to a java object
     * receiveType (in list) and excel sheet is one to one
     * <p>
     * 这种方法不推荐，还是推荐指定sheet页的方法。
     * 其一，是出错的概率比较高。
     * 其二，是获取结果相对繁琐。
     * 其三，这个方法间接调用了我推荐的，指定sheet页的方法。
     *
     * @param filepath
     * @param receiveType
     * @return
     * @throws Exception
     */
    public static <K> Map<String, List<K>> read(String filepath, List<Class<K>> receiveType) throws Exception {
        if (filepath == null || filepath.trim().isEmpty() || receiveType == null || receiveType.size() <= 0) {
            return null;
        }
        Map<String, List<K>> results = new HashMap<>(receiveType.size());
        for (Integer i = 0; i < receiveType.size(); i++) {
            if (receiveType.get(i) == null) {//这里实际上需要做的还有一些，比如排除class对象是接口，数组等等，暂时不考虑了。。
                continue;
            }
            List list = read(filepath, receiveType.get(i), i);
            results.put(i.toString(), list);
        }
        return results;
    }


    /**
     * write a excel file
     * if this filepath exists,we add a sheet in this  file!
     *
     * @param content sheet页里面具体的内容
     * @param filepath
     * @param sheetName sheet页的名称
     * @param <T>
     * @throws Exception
     */
    public static <T> void write(List<T> content, String filepath, String sheetName) throws Exception {
        if (content == null || content.size() <= 0 || filepath == null || filepath.trim().isEmpty()) {
            return;
        }

        String fileMark = filepath.substring(filepath.lastIndexOf("."));
        Class cls = content.get(0).getClass();
        Field[] fields = cls.getDeclaredFields();
        int length = fields.length;
        int size = content.size();

        ByteArrayOutputStream os = new ByteArrayOutputStream();
        File file = new File(filepath);//Excel文件生成后存储的位置。
        OutputStream fos = null;


        //如果存在了，那我们就继续往里面写内容，只是再另写一个sheet页，而不是追加内容！
        if (file.exists()) {
            Workbook workbook = readExcel(filepath);
            int sheetN = workbook.getNumberOfSheets();
            if (sheetName == null || sheetName.trim().isEmpty()) {
                sheetName = "sheet" + sheetN;
            }
            Sheet sheet = workbook.createSheet(sheetName);
            Row row = sheet.createRow(0);

            setContent(length, fields, row, size, sheet, content);

            writeExcel(workbook, os, fos, file);
            return;
        }


        if (sheetName == null || sheetName.trim().isEmpty()) {
            sheetName = "sheet1";
        }
        switch (fileMark) {
            case ".xls":
                //创建一个Excel
                HSSFWorkbook wb = new HSSFWorkbook();
                //创建一个工作簿sheet
                HSSFSheet sheet = wb.createSheet(sheetName);

                //创建sheet的 第一行
                HSSFRow rowTitle = sheet.createRow(0);

                setContent(length, fields, rowTitle, size, sheet, content);

                writeExcel(wb, os, fos, file);

                break;
            case ".xlsx":
                //创建一个Excel
                XSSFWorkbook xwb = new XSSFWorkbook();

                //创建一个工作簿sheet
                XSSFSheet xsheet = xwb.createSheet(sheetName);

                //创建sheet的 第一行
                XSSFRow rowTitlex = xsheet.createRow(0);

                setContent(length, fields, rowTitlex, size, xsheet, content);

                writeExcel(xwb, os, fos, file);

                break;
            default:
                throw new Exception("no this format");
        }


    }


    //写单元格的内容
    public static void setContent(int length, Field[] fields, Row rowTitle, int size, Sheet sheet, List content) throws Exception {
        //初始化第一行的内容
        for (int f = 0; f < length; f++) {
            String fieldName = fields[f].getName();
            System.out.println(fieldName);

            rowTitle.createCell(f);
            rowTitle.getCell(f).setCellValue(fieldName);
        }

        //设置单元格的具体内容
        for (int k = 0; k < size; k++) {
            //加一行
            Row rowAdd = sheet.createRow(k + 1);
            for (int f = 0; f < length; f++) {
                //设置属性可见
                fields[f].setAccessible(true);
                //由对象获取属性的值
                Object value = fields[f].get(content.get(k));
                //获取属性的类型
                String type = fields[f].getType().getSimpleName();
                //创建一个单元格
                Cell cels = rowAdd.createCell(f);
                if (value == null) {
                    continue;
                }
                //根据类型设置单元格数值
                setCellValue(type, rowAdd, f, value);
            }
        }

    }


    //写整个的Excel表
    private static void writeExcel(Workbook xwb, ByteArrayOutputStream os, OutputStream fos
            , File file) throws Exception {
        xwb.write(os);
        try {
            fos = new FileOutputStream(file);
            xwb.write(fos);
        } finally {
            os.close();
            if (fos != null) {
                fos.close();
            }
        }
    }

    //设置单元格的值
    private static void setCellValue(String type, Row rowAdd, int f, Object value) {
        switch (type) {
            case "String":
                rowAdd.getCell(f).setCellType(CellType.STRING);
                rowAdd.getCell(f).setCellValue(value.toString());
                break;
            case "Integer":
            case "int":
                rowAdd.getCell(f).setCellValue((int) value);
                break;
            case "Long":
            case "long":
                rowAdd.getCell(f).setCellValue((long) value);
                break;
            case "Double":
            case "double":
                rowAdd.getCell(f).setCellValue((Double) value);
                break;
            case "Boolean":
            case "boolean":
                rowAdd.getCell(f).setCellType(CellType.BOOLEAN);
                rowAdd.getCell(f).setCellValue((Boolean) value);
                break;
            case "Date":
                String date = new SimpleDateFormat("yyyy-MM-dd").format((Date) value);
                rowAdd.getCell(f).setCellValue(date);
                break;
            default:
        }
    }


    //读取excel，获取工作空间对象
    private static Workbook readExcel(String filePath) throws Exception {
//        Workbook wb = null;
        if (filePath == null || filePath.trim().isEmpty()) {
            return null;
        }
        String fileMark = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if (".xls".equals(fileMark)) {
                return new HSSFWorkbook(is);
            } else if (".xlsx".equals(fileMark)) {
                return new XSSFWorkbook(is);
            } else {
                return null;
            }
        } finally {
            if (is != null) {
                is.close();
            }
        }
    }

    /**
     * 获取单元格具体的类型值
     *
     * @param cell
     * @param valueType
     * @param <W>
     * @return
     */
    private static <W> W getCellFormatValue(Cell cell, Class<W> valueType) throws Exception {


        Object cellValue = null;
        if (cell == null) {
            cellValue = "";
            return (W) cellValue;
        }
        String returnType = valueType.getSimpleName();
        //判断cell类型
        switch (returnType) {
            case "String":
                cellValue = cell.getStringCellValue();
                break;
            case "Integer":
            case "int":
                cellValue = (int) cell.getNumericCellValue();
                break;
            case "Long":
            case "long":
                cellValue = (long) cell.getNumericCellValue();
                break;
            case "Double":
            case "double":
                cellValue = cell.getNumericCellValue();
                break;
            case "Boolean":
            case "boolean":
                if (CellType.STRING == cell.getCellType()) {
                    cellValue = Boolean.valueOf(cell.getStringCellValue());
                } else {
                    cellValue = cell.getBooleanCellValue();
                }
                break;
            case "Date":
                if (cell.getCellType() == CellType.STRING) {
                    cellValue = new SimpleDateFormat("yyyy-MM-dd").parse(cell.getStringCellValue());
                } else {
                    cellValue = cell.getDateCellValue();
                }
                break;
            default:
        }
        return (W) cellValue;

    }


}
