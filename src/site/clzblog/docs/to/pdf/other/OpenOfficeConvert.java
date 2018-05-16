package site.clzblog.docs.to.pdf.other;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

import java.io.File;
import java.io.IOException;

public class OpenOfficeConvert {

    public static void main(String[] args) {
        try {
            officeToPdf();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static int officeToPdf() throws IOException {
        String sourceFile = "/root/Documents/Test1111.doc";
        String destFile = "/root/Documents/Test1111.pdf";
        String OpenOffice_HOME = "";
        File inputFile = new File(sourceFile);
        if (!inputFile.exists()) {
            return -1;// 找不到源文件, 则返回-1
        }

        // 如果目标路径不存在, 则新建该路径
        File outputFile = new File(destFile);
        if (!outputFile.getParentFile().exists()) {
            outputFile.getParentFile().mkdirs();
        }

        //= "D:\\Program Files\\OpenOffice.org 3";//这里是OpenOffice的安装目录, 在我的项目中,为了便于拓展接口,没有直接写成这个样子,但是这样是绝对没问题的
        // 如果从文件中读取的URL地址最后一个字符不是 '\'，则添加'\'
       /* if (OpenOffice_HOME.charAt(OpenOffice_HOME.length() - 1) != '\\') {
            OpenOffice_HOME += "\\";
        }*/
        /*// 启动OpenOffice的服务
        String command = OpenOffice_HOME
                + "program\\soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;StarOffice.ServiceManager\" -nofirststartwizard";
        Process pro = Runtime.getRuntime().exec(command);*/
        // connect to an OpenOffice.org instance running on port 8100
        OpenOfficeConnection connection = new SocketOpenOfficeConnection("127.0.0.1", 8100);
        connection.connect();
        // convert
        DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
        //DocumentConverter converter = new StreamOpenOfficeDocumentConverter(connection);

        converter.convert(inputFile, outputFile);
        // close the connection
        connection.disconnect();
        // 关闭OpenOffice服务的进程
        //o.destroy();
        return 0;
    }
}
