package ExcelAPI;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 编辑excel
 *
 * @author JiangHaoDong
 */
public class Excel {
    //总行数
    private int totalRows = 0;
    //总列数
    private int totalCells = 0;
    //错误信息
    private String errorInfo;

    public Excel() {

    }

    //得到总行数
    public int getTotalRows() {
        return totalRows;
    }

    //得到总行数
    public int getTotalCells() {

        return totalCells;
    }

    //得到错误信息
    public String getErrorInfo() {

        return errorInfo;
    }

    /**
     * @描述：验证excel文件
     * @时间：2012-08-29 下午16:27:15
     * @参数：@param filePath　文件完整路径
     * @参数：@return
     * @返回值：boolean
     */

    public boolean validateExcel(String filePath) {
        /** 检查文件名是否为空或者是否是Excel格式的文件 */
        if (filePath == null || !(WDWUtil.isExcel2003(filePath) || WDWUtil.isExcel2007(filePath))) {

            errorInfo = "文件名不是excel格式";
            return false;
        }
/** 检查文件是否存在 */
        File file = new File(filePath);
        if (file == null || !file.exists()) {
            errorInfo = "文件不存在";
            return false;
        }
        return true;
    }

    /**
     * @描述：根据文件名读取excel文件
     * @时间：2012-08-29 下午16:27:15
     * @参数：@param filePath 文件完整路径
     * @参数：@return
     * @返回值：List
     */

    public List<Object> read(String filePath) {

        List<Object> dataLst = new ArrayList<Object>();
        InputStream is = null;
        try {
            /** 验证文件是否合法 */
            if (!validateExcel(filePath)) {

                System.out.println(errorInfo);
                return null;
            }

/** 判断文件的类型，是2003还是2007 */

            boolean isExcel2003 = true;
            if (WDWUtil.isExcel2007(filePath)) {

                isExcel2003 = false;
            }

/** 调用本类提供的根据流读取的方法 */

            File file = new File(filePath);
            is = new FileInputStream(file);
            dataLst = read(is, isExcel2003);
            is.close();
        } catch (Exception ex) {

            ex.printStackTrace();
        } finally {

            if (is != null) {

                try {

                    is.close();
                } catch (IOException e) {
                    is = null;
                    e.printStackTrace();
                }
            }
        }

        /** 返回最后读取的结果 */

        return dataLst;
    }

    /**
     * @描述：根据流读取Excel文件
     * @时间：2012-08-29 下午16:40:15
     * @参数：@param inputStream
     * @参数：@param isExcel2003
     * @参数：@return
     * @返回值：List
     */

    public List<Object> read(InputStream inputStream, boolean isExcel2003) {

        List<Object> dataLst = null;
        try {

            /** 根据版本选择创建Workbook的方式 */
            Workbook wb = null;
            if (isExcel2003) {
                wb = new HSSFWorkbook(inputStream);
            } else {
                wb = new XSSFWorkbook(inputStream);
            }

            dataLst = read(wb);
        } catch (IOException e) {

            e.printStackTrace();
        }

        return dataLst;
    }

    /**
     * @描述：根据流读取Excel文件，以列讀取
     * @时间：2012-08-29 下午16:40:15
     * @参数：@param inputStream
     * @参数：@param isExcel2003
     * @参数：@return
     * @返回值：List
     */

    public Map readByColumn(InputStream inputStream, boolean isExcel2003) {

        Map dataLst = null;
        try {

/** 根据版本选择创建Workbook的方式 */

            Workbook wb = null;
            if (isExcel2003) {

                wb = new HSSFWorkbook(inputStream);
            } else {

                wb = new XSSFWorkbook(inputStream);
            }

            dataLst = readByColumn(wb);
        } catch (IOException e) {

            e.printStackTrace();
        }

        return dataLst;
    }

    /**
     * @描述：读取数据，以行读取
     * @时间：2012-08-29 下午16:50:15
     * @参数：@param Workbook
     * @参数：@return
     * @返回值：List>
     */

    private List<Object> read(Workbook wb) {

        List<Object> dataLst = new ArrayList<Object>();
/** 得到第一个shell */

        Sheet sheet = wb.getSheetAt(0);
/** 得到Excel的行数 */

        this.totalRows = sheet.getPhysicalNumberOfRows();
/** 得到Excel的列数 */

        if (this.totalRows >= 1 && sheet.getRow(0) != null) {

            this.totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }

/** 循环Excel的行 */

        for (int r = 0; r < this.totalRows; r++) {

            Row row = sheet.getRow(r);
            if (row == null) {

                continue;
            }

            List rowLst = new ArrayList();
/** 循环Excel的列 */

            for (int c = 0; c < this.getTotalCells(); c++) {

                Cell cell = row.getCell(c);
                String cellValue = "";
                if (null != cell) {

// 以下是判断数据的类型

                    switch (cell.getCellType()) {

                        case HSSFCell.CELL_TYPE_NUMERIC: // 数字

                            cellValue = cell.getNumericCellValue() + "";
                            break;
                        case HSSFCell.CELL_TYPE_STRING: // 字符串

                            cellValue = cell.getStringCellValue();
                            break;
                        case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean

                            cellValue = cell.getBooleanCellValue() + "";
                            break;
                        case HSSFCell.CELL_TYPE_FORMULA: // 公式

                            cellValue = cell.getCellFormula() + "";
                            break;
                        case HSSFCell.CELL_TYPE_BLANK: // 空值

                            cellValue = "";
                            break;
                        case HSSFCell.CELL_TYPE_ERROR: // 故障

                            cellValue = "非法字符";
                            break;
                        default:

                            cellValue = "未知类型";
                            break;
                    }

                }

                rowLst.add(cellValue);
            }

/** 保存第r行的第c列 */

            dataLst.add(rowLst);
        }

        return dataLst;
    }

    /**
     * @描述：读取数据，以列读取
     * @时间：2012-08-29 下午16:50:15
     * @参数：@param Workbook
     * @参数：@return
     * @返回值：List>
     */

    public Map<Integer, Integer> readByColumn(Workbook wb) {

        Map<Integer, Integer> dataMap = new LinkedHashMap<>();
        /** 得到第一个shell */
        Sheet sheet = wb.getSheetAt(0);        /** 得到Excel的行数 */
        this.totalRows = sheet.getPhysicalNumberOfRows();
        /** 得到Excel的列数 */
        if (this.totalRows >= 1 && sheet.getRow(0) != null) {
            this.totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }

        //从第九列开始
        for (int lie = 0; lie <= totalCells; lie++) {
            Cell cell = sheet.getRow(1).getCell(lie);
            System.out.println(getCellValue(cell));            //设置为1，代表不需要删除，为0，则删除。
            dataMap.put(lie, 0);
        }

        /** 循环Excel的行(不包括第一行) */
        for (int r = 1; r < this.totalRows; r++) {

            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }

            /** 循环Excel的列 */
            for (int c = 0; c < this.getTotalCells(); c++) {
                Cell cell = row.getCell(c);
                if (null != cell && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
                    System.out.println(getCellValue(cell));
                    dataMap.put(c + 1, 1);
                }
            }
        }
        return dataMap;
    }

    /**
     * @描述：读取sheet数据，以列读取
     * @时间：2012-08-29 下午16:50:15
     * @参数：@param Workbook
     * @参数：@return
     * @返回值：List>
     */

    public Map readSheetByColumn(Sheet sheet) {

        Map dataMap = new LinkedHashMap<>();
/** 得到第一个shell */

/** 得到Excel的行数 */

        this.totalRows = sheet.getPhysicalNumberOfRows();
/** 得到Excel的列数 */

        if (this.totalRows >= 1 && sheet.getRow(0) != null) {

            this.totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }

//从第九列开始

        for (int lie = 9; lie <= totalCells; lie++) {

//设置为1，代表不需要删除，为0，则删除。

            dataMap.put(lie, 0);
        }

/** 循环Excel的行(不包括第一行) */

        for (int r = 1; r < this.totalRows; r++) {

            Row row = sheet.getRow(r);
            if (row == null) {

                continue;
            }

/** 循环Excel的列 */

            for (int c = 0; c < this.getTotalCells(); c++) {

                Cell cell = row.getCell(c);
                if (null != cell && cell.getCellType() != Cell.CELL_TYPE_BLANK) {

                    dataMap.put(c + 1, 1);
                }

            }

        }

        return dataMap;
    }

    /**
     * 删除列
     *
     * @param sheet
     * @param columnToDelete
     */

    public void deleteColumn(SXSSFSheet sheet, int columnToDelete) {

        for (int rId = 0; rId <= sheet.getLastRowNum(); rId++) {

            Row row = sheet.getRow(rId);
            for (int cID = columnToDelete; cID <= row.getLastCellNum(); cID++) {

                Cell cOld = row.getCell(cID);
                if (cOld != null) {

                    row.removeCell(cOld);
                }

                Cell cNext = row.getCell(cID + 1);
                if (cNext != null) {

                    Cell cNew = row.createCell(cID, cNext.getCellType());
                    cloneCell(cNew, cNext);
//Set the column width only on the first row.

//Other wise the second row will overwrite the original column width set previously.

                    if (rId == 0) {

                        sheet.setColumnWidth(cID, sheet.getColumnWidth(cID + 1));
                    }

                }

            }

        }

    }

    /**
     * 右边列左移
     *
     * @param cNew
     * @param cOld
     */

    private void cloneCell(Cell cNew, Cell cOld) {

        cNew.setCellComment(cOld.getCellComment());
        cNew.setCellStyle(cOld.getCellStyle());
        if (Cell.CELL_TYPE_BOOLEAN == cNew.getCellType()) {

            cNew.setCellValue(cOld.getBooleanCellValue());
        } else if (Cell.CELL_TYPE_NUMERIC == cNew.getCellType()) {

            cNew.setCellValue(cOld.getNumericCellValue());
        } else if (Cell.CELL_TYPE_STRING == cNew.getCellType()) {

            cNew.setCellValue(cOld.getStringCellValue());
        } else if (Cell.CELL_TYPE_ERROR == cNew.getCellType()) {

            cNew.setCellValue(cOld.getErrorCellValue());
        } else if (Cell.CELL_TYPE_FORMULA == cNew.getCellType()) {

            cNew.setCellValue(cOld.getCellFormula());
        }

    }

    /**
     * 读取第一行的cell数据
     */

    public List readFirstRow(SXSSFWorkbook wb) {

        System.out.println("进入readFirstRow");
        List list = new ArrayList<>();
        int totalColumn = 0;
        /** 得到第一个shell */

        Sheet sheet = wb.getSheetAt(0);
        /** 得到Excel的列数 */

        if (sheet.getRow(0) != null) {

            totalColumn = sheet.getRow(0).getPhysicalNumberOfCells();
            System.out.println("totalColumn:{}" + totalColumn);
            for (int lie = 8; lie < totalColumn; lie++) {

                System.out.println("lie:{}" + sheet.getRow(0).getCell(lie).getStringCellValue());
                list.add(sheet.getRow(0).getCell(lie).getStringCellValue());
            }

        }
        System.out.println();
        System.out.println("readFirst的大小：{}" + list.size());
        return list;
    }

    public static String getCellValue(Cell cell) {
        String cellValue = "";        // 以下是判断数据的类型
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC: // 数字
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    cellValue = sdf.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(cell.getNumericCellValue())).toString();
                } else {
                    DataFormatter dataFormatter = new DataFormatter();
                    cellValue = dataFormatter.formatCellValue(cell);
                }
                break;
            case Cell.CELL_TYPE_STRING: // 字符串
                cellValue = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_BOOLEAN: // Boolean
                cellValue = cell.getBooleanCellValue() + "";
                break;
            case Cell.CELL_TYPE_FORMULA: // 公式
                cellValue = cell.getCellFormula() + "";
                break;
            case Cell.CELL_TYPE_BLANK: // 空值
                cellValue = "";
                break;
            case Cell.CELL_TYPE_ERROR: // 故障
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }
        return cellValue;
    }

    public static void main(String[] args) {

        String filePath = "d://" + "dataFile.xls";
        File file = new File(filePath);
        Excel excel = new Excel();
        try {

            FileInputStream is = new FileInputStream(file);
            Workbook wb = WorkbookFactory.create(is);

            Map<Integer, Integer> stringIntegerMap = excel.readByColumn(wb);
            Iterator<Map.Entry<Integer, Integer>> it = stringIntegerMap.entrySet().iterator();
            while (it.hasNext()) {

                Map.Entry<Integer, Integer> entry = it.next();
                System.out.println("key= " + entry.getKey() + " and value= " + entry.getValue());
//                if (entry.getValue() == 0) {
//                    excel.deleteColumn(wb.getSheetAt(0), entry.getKey() - 1);//                }
            }

            FileOutputStream fileOut = new FileOutputStream(file);
            wb.write(fileOut);
            fileOut.close();
//excel.deleteColumn(sheet,1);
            System.out.println("删除成功");
        } catch (Exception e) {

            e.printStackTrace();
        }

    }

}

/**
 * @描述：工具类
 * @时间：2012-08-29 下午16:30:40
 */

class WDWUtil {

    /**
     * @描述：是否是2003的excel，返回true是2003
     * @时间：2012-08-29 下午16:29:11
     * @参数：@param filePath　文件完整路径
     * @参数：@return
     * @返回值：boolean
     */

    public static boolean isExcel2003(String filePath) {

        return filePath.matches("^.+\\.(?i)(xls)$");
    }

    /**
     * @描述：是否是2007的excel，返回true是2007
     * @时间：2012-08-29 下午16:28:20
     * @参数：@param filePath　文件完整路径
     * @参数：@return
     * @返回值：boolean
     */

    public static boolean isExcel2007(String filePath) {

        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }
}
