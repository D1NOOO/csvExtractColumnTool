
import java.io.*;
import java.util.Scanner;

/**
 * excel工具 v1.1
 * created by shenfeng on 2018.03.20
 * 读取Excel文件，输出某一列的数据到TXT文件中
 */
public class excelColumnTool {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        System.out.println("请输入被读文件的完整路径(包括文件名)：");
        String fileSource = scanner.next();

        System.out.println("文件是否包含标题行(Y/N)");
        String isTitle = scanner.next();
        while(!isTitle.toUpperCase().equals("Y") && !isTitle.toUpperCase().equals("N")){
                System.out.println("输入错误请重新输入！");
                isTitle = scanner.next();
        }

        System.out.println("请输入列数：");
        int column = scanner.nextInt();
        if(!(column>=1)){
            System.out.println("请输入正确的列数！");
        }

        System.out.println("读取文件中......");
        try {
            BufferedReader reader = new BufferedReader(new FileReader(fileSource));//换成你的文件名
            String txtName = null;
            if(isTitle.toUpperCase().equals("Y")){
                txtName = String.valueOf(System.currentTimeMillis());//文件名前缀
                reader.readLine();//第一行信息，为标题信息。需要标题N，不需要标题Y。
            }else {
                txtName = reader.readLine().split(",")[column-1];
            }
            String line = null;
            System.out.println("请输入输出路径：");
            String fileTarget = scanner.next();
            BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(new File(fileTarget+"\\"+txtName+"_result.txt")),
                    "UTF-8"));
            while((line=reader.readLine())!=null){
                if(column>line.split(",").length){
                    System.out.println("超出列数范围！");
                    System.exit(-1);
                }
                String item = line.split(",")[column-1];//CSV格式文件为逗号分隔符文件，这里根据逗号切分
                System.out.println(item);
                bw.write(item);
                bw.newLine();
            }
            bw.flush();
            bw.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
