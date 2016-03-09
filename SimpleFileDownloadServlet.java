package cn.com.genechem.download;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLDecoder;
import java.util.Hashtable;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.StringUtils;

import cn.com.genechem.util.FileUtils;

/**
 * 文件简单下载servlet, 缺省disable，不能调用<br>
 * 使用http://127.0.0.1/xx/simpledownload/base root下的目录结构/文件名<br>
 * 先调用SimpleFileDownloadServlet.initBaseRoot设置文件的base root目录或者调用enable()启用servlet
 */
@WebServlet("/simpledownload/*")
public class SimpleFileDownloadServlet extends HttpServlet{
    
    private static final long serialVersionUID = 2200171994814548056L;
    
    // 文件的根目录
    private static String base_folder = null;

    private static Hashtable<String, String> htAllowedIps = new Hashtable<String, String>();
    private static Hashtable<String, String> htAllowedDirs = new Hashtable<String, String> ();
    
    // servlet是否启用，缺省不启用
    private static boolean enbale = false;

    private static boolean checkIP = false;

    private static boolean checkDir = false;
    
    public static void initBaseRoot(String folder){
        File file = new File(folder);
        if(file!=null && !file.exists()){
            FileUtils.mkDir(file);
        }
        
        base_folder = folder;
        enbale = true;
    }

    private static boolean isIPAllowed (String ip) {
        if (StringUtils.isBlank(ip)) return false;
        if (StringUtils.equals(htAllowedIps.get(ip),"IP")) return true;
        for (String s:htAllowedIps.keySet()) {
            if (ip.startsWith(s)) return true;
        }
        return false;
    }

    private static boolean isDirAllowed (String path) {
        if (StringUtils.isBlank(path)) return false;
        for (String dir:htAllowedDirs.keySet()) {
            if (path.startsWith(dir)) return true;
        }
        return false;
    }
    
    public static void enbale(){
        enbale = true;
    }

    public static void enbaleIpCheck (String [] ipList) {
        checkIP = true;
        if (null == ipList) return;
        for (String str:ipList) {
            htAllowedIps.put(str,"IP");
        }
    }

    public static void enbaleDirCheck (String [] ipList) {
        checkDir = true;
        if (null == ipList) return;
        for (String str:ipList) {
            htAllowedDirs.put(str,"DIR");
        }
    }

    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException,IOException {
        if(!enbale)
            return;
        
        if(request.getServletPath().equals("/simpledownload")){
            downloadHandle(request, response);
        }
    }
    
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException,IOException {
        if(!enbale)
            return;
        
        if(request.getServletPath().equals("/simpledownload")){
            downloadHandle(request, response);
        }
    }
    
    private void downloadHandle(HttpServletRequest request, HttpServletResponse response){
        response.setCharacterEncoding("UTF-8");
        
        // 此处没做session检查，应该在filter集中处理
        
        try {
            String uri = URLDecoder.decode(request.getRequestURI(), "utf-8");
            String servlet = request.getServletPath();
            
            String filePath = null;
            if(StringUtils.isEmpty(base_folder)){
                filePath = uri.substring(uri.indexOf(servlet)+servlet.length());
            }
            else{
                filePath = base_folder + uri.substring(uri.indexOf(servlet)+servlet.length());
            }
            
            File file = new File(filePath);

            if (checkDir && !isDirAllowed(file.getCanonicalPath ())) {
                response.setContentType("text/html;charset=utf-8");
                response.getWriter().write("文件目录不允许访问！");
            }

            if (checkIP && !isIPAllowed(request.getRemoteAddr())) {
                response.setContentType("text/html;charset=utf-8");
                response.getWriter().write("客户端IP地址无访问权限！");
            }

            if (file != null && file.exists() && file.isFile() && file.canRead()) {
                InputStream fis = new BufferedInputStream(new FileInputStream(file));
                byte[] buffer = new byte[fis.available()];
                fis.read(buffer);
                fis.close();

                response.setContentType("application/octet-stream");
                response.getOutputStream().write(buffer);
            } else {
                System.out.println("file not exits: "+filePath);
                response.sendError(HttpServletResponse.SC_NOT_FOUND);
                return;
            }
        } catch (Exception e) {
            e.printStackTrace();
            response.setContentType("text/html;charset=utf-8");
            try {
                response.getWriter().write("服务器内部错误");
            } catch (IOException e1) {
                e1.printStackTrace();
            }
        }
    }
}
