package com.demo;

import com.aspose.words.Document;
import com.aspose.words.License;
import com.aspose.words.net.System.Data.DataRelation;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.net.System.Data.DataTable;

import java.io.FileInputStream;
import java.io.InputStream;

/**
 * @author bu.han
 */
public class AsposeWordFill {

    private static InputStream fileInput;
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        System.out.println("Start...");

        //第一步 加载授权文件license.xml
        //没有授权的情况下创建的word会出现水印
        License aposeLic = new License();
        try {
            ClassLoader loader = Thread.currentThread().getContextClassLoader();
            // 凭证文件
            InputStream license = new FileInputStream(loader.getResource("license.xml").getPath());
            aposeLic.setLicense(license);
            // 待处理的文件
            fileInput = new FileInputStream(loader.getResource("xiaoguo123.docx").getPath());
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }

        //模板文件和将要创建的新文件
        //String template = "xiaoguo123.docx";    //可以是doc或docx
        String destdoc = "xiaoguo123_new.docx"; //可以是doc或docx
        Document doc;
        try{
            //第二步 读取word模板文件，可以是.doc或者.docx
            doc = new Document( fileInput );

            //第三步 向模板中填充数据
            //主要调用aspose.words的邮件合并接口MailMerge
            //3.1 填充单个文本域
            String[] Flds = new String[]{"Title", "Name", "URL", "Note"}; //文本域
            String Name = "小郭软件";
            String URL = "http://xiaoguo123.com";
            String Note = "分享绿色简单易用的小软件";
            //值
            Object[] Vals = new Object[]{"小郭软件@2016", Name, URL, Note };
            //调用接口
            doc.getMailMerge().execute(Flds, Vals);

            //3.2 填充单层循环的表格
            //网站访问量表格
            DataTable visitTb = new DataTable("Visit");
            visitTb.getColumns().add("Date"); //0 增加三个列 日期
            visitTb.getColumns().add("IP");   //1 IP访问数量
            visitTb.getColumns().add("PV");   //2 页面浏览量
            //向表格中填充数据
            for(int i=1; i<3; i++){
                DataRow row = visitTb.newRow(); //新增一行
                row.set(0, "2016年2月"+i+"日"); //根据列顺序填入数值
                row.set(1, i*300);
                row.set(2, i*400);
                //加入此行数据
                visitTb.getRows().add( row );
            }
            //对于无数据的情况，增加一行空记录
            if( visitTb.getRows().getCount() == 0 ){
                DataRow row = visitTb.newRow();
                visitTb.getRows().add( row );
            }
            //调用接口
            doc.getMailMerge().executeWithRegions( visitTb );

            //3.3 填充具有两层循环的表格
            //需要定义两个数据表格，且两者之间通过某列关联起来
            //用户
            DataTable userTb = new DataTable("User");
            //用户名称
            userTb.getColumns().add("Name");
            userTb.getColumns().add("RegDate");
            //用户信息
            DataTable infoTb = new DataTable("Info");
            //用户名称 通过此列和上个表User关联
            infoTb.getColumns().add("Name");
            infoTb.getColumns().add("Date");
            infoTb.getColumns().add("Time");

            //3.3.1 填充用户信息
            for(int i=1; i<4; i++){
                DataRow row = userTb.newRow();
                row.set(0, "User"+i );
                row.set(1, "2015年3月"+i+"日");
                userTb.getRows().add( row );
            }
            //3.3.2 填充详细信息
            for(int i=1; i<6; i++){
                for(int j=1; j<5; j++){
                    DataRow row = infoTb.newRow();
                    row.set(0, "User"+i );
                    row.set(1, "2016年1月"+j+"日");
                    row.set(2, j*2*i );
                    infoTb.getRows().add( row );
                }
            }

            //3.3.3 将 User 和 Info 关联起来
            DataSet userSet = new DataSet();
            userSet.getTables().add( userTb );
            userSet.getTables().add( infoTb );
            String[] contCols = {"Name"};
            String[] lstCols = {"Name"};
            userSet.getRelations().add( new DataRelation("UserInfo", userTb, infoTb, contCols, lstCols) );
            //调用接口
            doc.getMailMerge().executeWithRegions(userSet);

            //第四步 保存新word文档
            doc.save( destdoc );
            System.out.println("End...");
        }catch (Exception e) {
            System.out.println(e.getMessage());
        }

    }
}
