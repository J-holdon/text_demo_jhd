package ExcelAPI;


import model.GoodsPrice;
import enums.TableType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.UUID;
import java.util.concurrent.atomic.AtomicInteger;

public class downloadExcel {

    public static void main(String[] args) throws IOException {
        List<LinkedHashMap<String, Object>> linkedHashMaps = new ArrayList<LinkedHashMap<String, Object>>();
        for (int i = 0; i < 10; i++) {
            LinkedHashMap<String, Object> linkedHashMap = new LinkedHashMap<>();
            GoodsPrice goodsPrice = new GoodsPrice();
            goodsPrice.setBidderName("测试投标人" + i);
            goodsPrice.setCompanyName("测试招标人" + i);
            goodsPrice.setGoodsName(UUID.randomUUID().toString());
            goodsPrice.setGoodsCode("测试物资编码");
            goodsPrice.setPrice(BigDecimal.valueOf(100 + i));
            goodsPrice.setResponsePrice(BigDecimal.valueOf(i * i));
            linkedHashMap.put("供应商名称", goodsPrice.getBidderName());
            linkedHashMap.put("采购方名称", goodsPrice.getCompanyName());
            linkedHashMap.put("物资名称", goodsPrice.getGoodsName());
            linkedHashMap.put("物资编码", goodsPrice.getGoodsCode());
            linkedHashMap.put("单价", goodsPrice.getPrice());
            linkedHashMap.put("响应单价", goodsPrice.getResponsePrice());
            linkedHashMaps.add(linkedHashMap);
        }
        ExportExecl(linkedHashMaps, "测试excel");
        addExcel(linkedHashMaps, "测试excel");
    }

    /**
     * 导出excel
     * @param dataList 数据
     * @param name 文件名称
     */
    public static void ExportExecl(List<LinkedHashMap<String,Object>> dataList, String name){
        LinkedHashMap<String,Object> titleList = dataList.get(0);
        // 创建HSSFWorkbook对象
        HSSFWorkbook wb = new HSSFWorkbook();
        // 创建HSSFSheet对象
        HSSFSheet sheet = wb.createSheet(name);
        HSSFCellStyle cellStyle = wb.createCellStyle();
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
        cellStyle.setWrapText(true);//自动换行
        // -----------------------------数据填充 ---------------------------------
        // 设置标题
        HSSFRow titleRow = sheet.createRow(0);
        HSSFCell titleCell=titleRow.createCell(0);
        titleCell.setCellValue(name);
        titleCell.setCellStyle(cellStyle);
        //合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,titleList.size()-1));
        // 设置列名称
        HSSFRow cellTitleRow = sheet.createRow(1);
        AtomicInteger jt = new AtomicInteger();
        LinkedHashMap<String ,Object> tem = dataList.get(0);
        tem.forEach((k,v)-> {
            HSSFCell cell=cellTitleRow.createCell(jt.get());
            cell.setCellValue(k);
            cell.setCellStyle(cellStyle);
            jt.getAndIncrement();
        });
        // 填充sql结果
        for (int i = 0; i < dataList.size(); i++){
            HSSFRow row = sheet.createRow(i+2);
            AtomicInteger j = new AtomicInteger();
            LinkedHashMap<String ,Object> ltem = dataList.get(i);
            ltem.forEach((k,v)-> {
                HSSFCell cell=row.createCell(j.get());
                cell.setCellValue(v.toString());
                cell.setCellStyle(cellStyle);
                j.getAndIncrement();
            });
        }
        // -----------------------------------------------------------------------
        // 输出Excel文件
        try {
            FileOutputStream output = new FileOutputStream("d:\\"+name+".xls");
            wb.write(output);
            output.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 改变表格样式
     */
    private HSSFCellStyle changeStyle(HSSFCellStyle cellStyle, TableType type){
        switch (type){
            case T_HEED:
                //.....todo 表格不同部分的样式
                break;
            case NORMAL:
                //.....todo 表格不同部分的样式
                break;
            case T_BODY:
                //.....todo 表格不同部分的样式
                cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
                cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
                cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
                cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
                cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
                cellStyle.setWrapText(true);//自动换行
                break;
            case T_FOOT:
                //.....todo 表格不同部分的样式
                break;
            default:
                break;
        }
        return cellStyle;
    }


    /**
     * 追加到已有excel
     * @param dataList 数据
     * @param name 文件名
     */

    public static void addExcel(List<LinkedHashMap<String,Object>> dataList, String name) throws IOException {
        FileInputStream fileInputStream=new FileInputStream("d://"+name+".xls");  //获取d://test.xls,建立数据的输入通道
        POIFSFileSystem poifsFileSystem=new POIFSFileSystem(fileInputStream);  //使用POI提供的方法得到excel的信息
        HSSFWorkbook Workbook=new HSSFWorkbook(poifsFileSystem);//得到文档对象
        HSSFSheet sheet=Workbook.getSheet(name);  //根据name获取sheet表
        HSSFCellStyle cellStyle = Workbook.createCellStyle();
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
        //合并单元格
        //参数1：起始行 参数2：终止行 参数3：起始列 参数4：终止列
        CellRangeAddress region1 = new CellRangeAddress(sheet.getLastRowNum() + dataList.size() + 1, sheet.getLastRowNum() + dataList.size() + 1, (short) 0, (short) 11);
        sheet.addMergedRegion(region1);
        // HSSFRow row=sheet.getRow(0);  //获取第一行
        System.out.println("最后一行的行号 :"+sheet.getLastRowNum() + 1);  //分别得到最后一行的行号，和第3条记录的最后一个单元格
//        System.out.println("最后一个单元格 :"+row.getLastCellNum());  //分别得到最后一行的行号，和第3条记录的最后一个单元格
        // HSSFRow startRow=sheet.createRow((short)(sheet.getLastRowNum()+1)); // 追加开始行
        // -----------------追加数据-------------------
        int start = sheet.getLastRowNum() + 1; //插入数据开始行
        for (int i = 0; i < dataList.size(); i++){
            HSSFRow startRow = sheet.createRow(i+start);
            AtomicInteger j = new AtomicInteger();
            LinkedHashMap<String ,Object> ltem = dataList.get(i);
            ltem.forEach((k,v)-> {
                HSSFCell cell=startRow.createCell(j.get());
                cell.setCellValue(v.toString());
                cell.setCellStyle(cellStyle);
                j.getAndIncrement();
            });
        }
        // 输出Excel文件
        FileOutputStream out=new FileOutputStream("d://"+name+".xls");  //向d://test.xls中写数据
        out.flush();
        Workbook.write(out);
        out.close();
    }
}
