package com.test.word.wordpdf.model;

import org.apache.commons.io.IOUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.*;

/**
 * linux或windows命令执行
 */
public class CommandExecute {
    //    public static void main(String[] args) {
//        CommandExecute obj = new CommandExecute();
//        String domainName = "www.baidu.com";
//        //in mac oxs
//        String command = "ping " + domainName;
//        //in windows
//        //String command = "ping -n 3 " + domainName;
//        String output = obj.executeCommand(command);
//        System.out.println(output);
//    }
    public static String executeCommand(String command) {
        StringBuffer output = new StringBuffer();
        Process p;
        InputStreamReader inputStreamReader = null;
        BufferedReader reader = null;
        try {
            p = Runtime.getRuntime().exec(command);
            p.waitFor();
            inputStreamReader = new InputStreamReader(p.getInputStream(), "UTF-8");
            reader = new BufferedReader(inputStreamReader);
            String line = "";
            while ((line = reader.readLine()) != null) {
                output.append(line + "\n");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InterruptedException e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(reader);
            IOUtils.closeQuietly(inputStreamReader);
        }
        System.out.println(output.toString());
        return output.toString();
    }

    public static void findPdf(HttpServletResponse response, String targetPath) throws IOException {
        response.setContentType("application/pdf");
        response.setCharacterEncoding("utf-8");
        FileInputStream in = new FileInputStream(new File(targetPath));
        OutputStream out = response.getOutputStream();
        byte[] b = new byte[512];
        while ((in.read(b)) != -1) {
            out.write(b);
        }
        out.flush();
        in.close();
        out.close();
    }
}