import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Scanner;

/**
 * Created by azure on 2015/5/30 0030.
 */
public class MergeData
{
    private static Logger logger = LogManager.getLogger(MergeData.class);

    private static Path getPath(String basePath, String fileName)
    {
        System.out.print("输入" + fileName + "文件名(或全路径): ");
        Scanner scanner = new Scanner(System.in);
        String p = scanner.next();
        Path path1 = null;
        if(("exit").equals(p))
        {
            logger.info("退出系统");
            System.exit(0);
        }
        try
        {
            while (true)
            {
                if (p.indexOf(":") == -1)
                {
                    path1 = Paths.get(basePath + p);
                } else
                {
                    path1 = Paths.get(p);
                }
                if (!path1.toFile().exists())
                {
                    System.out.print("文件不存在，请重新输入" + fileName + "文件名(或全路径): ");
                    p = scanner.next();
                } else
                {
                    logger.info("文件存在，文件路径：" + path1.toRealPath());
                    break;
                }
            }
        } catch (IOException e)
        {
            e.printStackTrace();
        }
        return path1;
    }

    private static String getInput(String prom)
    {
        System.out.print(prom + ": ");
        Scanner scanner = new Scanner(System.in);
        String p = scanner.next();
        if(("exit").equals(p))
        {
            System.exit(0);
        }
        return p;
    }

    private static String getInput(String[] title, String prom)
    {
        prom = "请从" + Arrays.toString(title) + "中选择 " + prom + " : ";
        System.out.print(prom);
        Scanner scanner = new Scanner(System.in);
        String p = scanner.next();
        if(("exit").equals(p))
        {
            System.exit(0);
        }
        while (true)
        {
            int i = 0;
            for (; i < title.length; i++)
            {
                if (p.equals(title[i]))
                {
                    break;
                }
            }
            if (i == title.length)
            {
                System.out.print("不存在该值，" + prom.replace("请", "请重新"));
                p = scanner.next();
            } else if (i < title.length)
            {
                break;
            }
        }
        return p;
    }


    /**
     * args[0]:Excel_1 path
     * args[1]:Excel_2 path
     * args[2]:0,1
     * args[3]:merge column
     *
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception
    {
        String path = MergeData.class.getProtectionDomain().getCodeSource().getLocation().toString();
        if (path.endsWith(".jar"))
        {
            path = path.substring(0, path.lastIndexOf("/") + 1);
        }
        path = path.replace("file:/", "");
        while (true)
        {
            Path path1 = getPath(path, "第一个");
            Path path2 = getPath(path, "第二个");
            HandleData handleData = new HandleData();
            handleData.path = path1;
            handleData.resultPathStr = path1.toRealPath().toString().replace(".xls", "_result.xls");
            InputStream inputStream = new FileInputStream(path1.toRealPath().toString());
            String[] title1 = handleData.title = handleData.readExcelTitle(inputStream);
            handleData.data1 = new ArrayList<>();
            handleData.readExcelContent(handleData.data1);
            handleData.path = path2;
            inputStream = new FileInputStream(path2.toRealPath().toString());
            String[] title2 = handleData.title = handleData.readExcelTitle(inputStream);
            handleData.column1 = getInput(title2, "合并列名");
            handleData.data2 = new ArrayList<>();
            handleData.readExcelContent(handleData.data2);
            handleData.updateUID(handleData.data1, "year", "id");
            handleData.updateUID(handleData.data2, "year", "id");
            handleData.resultData = new HashMap<>();
            handleData.mergeData(handleData.resultData, handleData.data1);
            handleData.mergeData(handleData.resultData, handleData.data2, handleData.column1);
            String[] title = new String[title1.length + 1];
            title = Arrays.copyOf(title1, title.length);
            title[title.length - 1] = handleData.column1;
            handleData.title = title;
            HandleData.logger.info("title:" + Arrays.toString(title));
            handleData.data2 = null;
            if (getInput("是否需要排序(yes/no)").toLowerCase().indexOf("y") != -1)
            {
                handleData.data2 = new ArrayList<>();
                String sortName = getInput(title, "排序名");

                handleData.sortData(handleData.resultData, handleData.data2, sortName);
            }
            handleData.writeResultData(handleData.data2);
            if (getInput("本次处理完成，是否继续处理新的数据(yes/no)").toLowerCase().indexOf("y") == -1)
                break;
        }

    }

}
