package test;

import org.lszjaf.utils.excel.ExcelUtil;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelTest {


    public static void main(String[] args) throws Exception {
//        String filePath = "D:\\test.xls";
        String filePath = "D:\\testwws.xlsx";
        System.out.println();
        List listst = new ArrayList();
        listst.add(ExcelBean.class);
//        listst.add(ExcelBeanTwo.class);
//        Map<String,List<ExcelBean> > list = ExcelUtil.read(filePath,listst);
//        System.out.println(list);

        List listss = new ArrayList();
        ExcelBean excelBean = new ExcelBean();
        excelBean.setBalance(12);
        excelBean.setName("32143214");
        excelBean.setAge(43);
        excelBean.setFlag(true);
        excelBean.setIdCard(988999);
        excelBean.setUpdateTime(new Date());
        listss.add(excelBean);

        ExcelBean excelBean1 = new ExcelBean();
        excelBean1.setBalance(32);
        excelBean1.setName("rewqr314324");
        excelBean1.setAge(323);
        excelBean1.setFlag(false);
        excelBean1.setIdCard(4444);
        excelBean1.setUpdateTime(new Date());
        listss.add(excelBean1);
//        String filePath = "D:\\testww.xls";
        ExcelUtil.write(listss,filePath,null);
    }


}
