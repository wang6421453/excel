package com.wl.excel;

import com.wl.excel.manager.ExcelContext;
import com.wl.excel.manager.ExcelManager;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel测试类
 * @author  by wl
 * @date 2017/11/8.
 */
public class ExcelTest {

    @Test
    public void test() throws Exception{
        ExcelManager excelManager = ExcelManager.getInstance();
        ExcelContext excelContext = excelManager.generateExcelContext("E:\\excel\\test.xlsx", ImportVo.class);
        List<ImportVo> lstImportVo = (List<ImportVo>)excelManager.importExcel(excelContext);
        for (ImportVo v : lstImportVo) {
            System.out.println(v.getUserName() + "  " + v.getAge() + "  " + v.getSex());
        }

        List<ExportVo> lstExport = new ArrayList<ExportVo>();
        for(int i = 0; i < 10; i++){
            ExportVo vo = new ExportVo();
            vo.setUserName("w" + i);
            vo.setAge(i);
            vo.setSex("男");
            lstExport.add(vo);
        }

        ExcelContext excelContextExport =  excelManager.generateExcelContext("E:\\excel\\test_export.xlsx", ExportVo.class);
        excelManager.exportExcel(excelContextExport, lstExport);
    }

}
