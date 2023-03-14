package com.feiyizhan.excel.export.utils;

import lombok.extern.log4j.Log4j2;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel 工具类
 * @author 徐明龙 XuMingLong 2023-03-02
 */
@Log4j2
public class ExcelUtil {
    private ExcelUtil() {
        throw new IllegalStateException("Utility class");
    }


    /**
     * 行复制规则
     * @author 徐明龙 XuMingLong 2023-03-01
     */
    private static final CellCopyPolicy DEFAULT_ROW_COPY_POLICY = new CellCopyPolicy.Builder().build();


    /**
     * 移动指定的选择区间行区间到指定位置，并删除移动后原地的行，类似于Excel操作中的剪切复制
     * @author 徐明龙 XuMingLong 2023-03-02
     * @param sheet 待处理的Sheet
     * @param srcBeginRow 待移动行区域的开始行号
     * @param srcEndRow  待移动行区域的结束行号
     * @param destBeginRow 目标区域的开始行号
     * @return void
     */
    public static void moveRows(XSSFSheet sheet,int srcBeginRow,int srcEndRow,int destBeginRow){
        //先复制到目的位置
        sheet.copyRows(srcBeginRow, srcEndRow, destBeginRow, DEFAULT_ROW_COPY_POLICY);
        int movedRows = srcEndRow - srcBeginRow+ 1;
        int destEndRow = destBeginRow + movedRows;
        //清理之前残留的合并单元格
        List<Integer> removeMergedRegion = new ArrayList<>();
        for(int i=sheet.getMergedRegions().size()-1;i>=0;i--){
            CellRangeAddress address = sheet.getMergedRegions().get(i);
            if(srcBeginRow < destBeginRow){
                //如果合并单元格在复制前的位置，则删除
                if(address.getFirstRow() >= srcBeginRow && address.getFirstRow()<destBeginRow ){
                    removeMergedRegion.add(i);
                }
            }else{
                //如果合并单元格在复制后位置，则删除
                if(address.getFirstRow() >= destEndRow){
                    removeMergedRegion.add(i);
                }
            }

        }
        //执行清理合并的单元格
        for(Integer i:removeMergedRegion){
            sheet.removeMergedRegion(i);
        }
        //删除移动后原地的行
        for (int i = srcBeginRow; i <= srcEndRow; i++){
            Row row = sheet.getRow(i);
            if (row != null){
                //删除复制后残留的行
                sheet.removeRow(row);
            }
        }

    }


    /**
     * 在指定的行后面插入指定的空行数，并支持设置插入的空行的样式参考指定行样式(使用ShiftRows方式）
     * @author 徐明龙 XuMingLong 2023-03-14
     * @param sheet 待操作的Sheet
     * @param afterRowNo 指定的行的行号，将在这个行后面插入空行。
     * @param insertedRowsNumber 待插入的空行数量
     * @param copyStyleRowNo 复制的行样式的行号，-1 表示只创建空白行，不复制样式。
     * @return void
     */
    public static void insertRowsAfterRowWithShiftRows(XSSFSheet sheet,int afterRowNo,int insertedRowsNumber,int copyStyleRowNo){
        int moveBeginRow = afterRowNo+1;
        //如果移动的开始行不存在，则创建一个空行，否则复制的时候会丢掉不存在的行
        if(sheet.getRow(moveBeginRow)==null){
            sheet.createRow(moveBeginRow);
        }
        int lastRowNo = sheet.getLastRowNum();
        //计算移动的行数
        int movedNo = lastRowNo-moveBeginRow;
        //移动待添加的行
        sheet.shiftRows(moveBeginRow,lastRowNo,insertedRowsNumber);
        //如果参考的行在插入行后面，则说明参考的行在移动行后有变动，需要重新计算
        if(copyStyleRowNo > afterRowNo){
            //计算移动后的参考行号
            copyStyleRowNo = copyStyleRowNo+movedNo;
        }
        //末尾部分的最终开始行
        int finalBeginRowNo = insertedRowsNumber+afterRowNo;
        //在空出位置插入指定行数的新行并复制参考行的样式
        createBlankRows(sheet,moveBeginRow,finalBeginRowNo,copyStyleRowNo);
    }

    /**
     * 在指定的行后面插入指定的空行数，并支持设置插入的空行的样式参考指定行样式
     * @author 徐明龙 XuMingLong 2023-03-02
     * @param sheet 待操作的Sheet
     * @param afterRowNo 指定的行的行号，将在这个行后面插入空行。
     * @param insertedRowsNumber 待插入的空行数量
     * @param copyStyleRowNo 复制的行样式的行号，-1 表示只创建空白行，不复制样式。
     * @return void
     */
    public static void insertRowsAfterRow(XSSFSheet sheet,int afterRowNo,int insertedRowsNumber,int copyStyleRowNo){
        int moveBeginRow = afterRowNo+1;
        //如果移动的开始行不存在，则创建一个空行，否则复制的时候会丢掉不存在的行
        if(sheet.getRow(moveBeginRow)==null){
            sheet.createRow(moveBeginRow);
        }
        int lastRowNo = sheet.getLastRowNum();

        //计算移动的行数
        int movedNo = lastRowNo-moveBeginRow+1;
        //临时区的开始行
        int tempBeginRowNo = lastRowNo+insertedRowsNumber+1;
        //临时区的结束行
        int tempLastRowNo = tempBeginRowNo+movedNo;
        //末尾部分的最终开始行
        int finalBeginRowNo = insertedRowsNumber+afterRowNo+1;
        //移动开始行以后的所有行到临时区的开始行
        moveRows(sheet,moveBeginRow,lastRowNo,tempBeginRowNo);
        //如果参考的行在插入行后面，则说明参考的行在移动行后有变动，需要重新计算
        if(copyStyleRowNo > afterRowNo){
            //计算移动后的参考行号
            copyStyleRowNo = copyStyleRowNo+movedNo;
        }
        //在空出位置插入指定行数的新行并复制参考行的样式
        createBlankRows(sheet,moveBeginRow,finalBeginRowNo,copyStyleRowNo);
        //将移动的行从临时开始行回移新增好的行的末尾
        /*
          为什么不直接第一次移动的预期的末尾呢？是因为POX只支持Copy行的操作，而copy操作不是完整的区间覆盖。
          如果源区间内带有格式，并且目的区间和源区间有重叠话，会导致无法copy后源区间和目的区间混合在一起，无法清理。
          因此实现移动行区间的操作时，需要先将源行区间复制到一个临时的区域，然后清理掉源行区间的行和样式，再插入需要新增的行
          ，最后再将被移动的行区间移动回新的行末尾。
         */
        moveRows(sheet,tempBeginRowNo,tempLastRowNo,finalBeginRowNo);
    }

    /**
     * 创建多个空白行，支持设置空白行的参考样式行
     * @author 徐明龙 XuMingLong 2023-03-02
     * @param sheet 待操作的Sheet
     * @param beginRowNo 创建行的起始行号（含）
     * @param endRowNo 创建行的截至行号（不含）
     * @param copyStyleRowNo 复制的行样式的行号，-1 表示只创建空白行，不复制样式。
     * @return void
     */
    public static void createBlankRows(XSSFSheet sheet,int beginRowNo,int endRowNo,int copyStyleRowNo){
        Row copyStyleRow = sheet.getRow(copyStyleRowNo);
        for (int i = beginRowNo; i <endRowNo; i++){
            Row row = sheet.getRow(i);
            if (row != null){
                sheet.removeRow(row);
            }
            //创建空白的行
            sheet.createRow(i);
            if(copyStyleRow!=null){
                //复制第一行的样式到当前行
                sheet.copyRows(copyStyleRowNo,copyStyleRowNo, i, DEFAULT_ROW_COPY_POLICY);
            }
        }
    }

    /**
     * 根据指定单元格创建不共享的单元格样式
     * @author 徐明龙 XuMingLong 2023-03-02
     * @param cell
     * @return org.apache.poi.xssf.usermodel.XSSFCellStyle
     */
    public static XSSFCellStyle createNotShareCellStyle(XSSFCell cell){
        XSSFCellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        cellStyle.cloneStyleFrom(cell.getCellStyle());
        return cellStyle;
    }

    /**
     * 根据指定单元格创建共享的单元格样式
     * @author 徐明龙 XuMingLong 2023-03-02
     * @param cell
     * @return org.apache.poi.xssf.usermodel.XSSFCellStyle
     */
    public static XSSFCellStyle createShareCellStyle(XSSFCell cell){
        return cell.getCellStyle();
    }
}
