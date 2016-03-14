package cn.com.genechem.doc;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class DocResult {
    URL template;
    DocTemplate templateDocx;
    DocModel rootBean;

    public DocResult(URL template,DocModel model) throws Exception {
        this.template = template; 
        rootBean = model;
        //outputPath = output;
        InputStream in = template.openStream();
        templateDocx = new DocTemplate(in);
        in.close();
        generate();
    }

    void generate() throws Exception {
        // templateDocx.getDocument().dump();
        replaceBookmarks(rootBean,templateDocx.getBookmarkRootNode());
        templateDocx.cleanRemovedElems();

    }
    public void save(String outputPath) throws Exception{
        FileOutputStream out = new FileOutputStream(outputPath);
        templateDocx.write(out);
        out.close();
    }
    public void save(OutputStream stream) throws Exception{
        templateDocx.write(stream);
    }
    void replaceBookmarks(Object bean, BookmarkTreeNode bookmarkNode) {
        Object result = getProperty(bean, bookmarkNode);
        
        List<BookmarkTreeNode> children = bookmarkNode.children;
        //Bookmark bookmark = bookmarkNode.bookmark;
        
        if(result!=null){
            bookmarkNode.reset();
        }
        
        if(result==null){
            bookmarkNode.remove();
            for(BookmarkTreeNode child:children){
                replaceBookmarks(null,child);
            }
        } else if(result instanceof ElemCustTableHead){
            ElemCustTableHead th = (ElemCustTableHead)result;
            bookmarkNode.bookmark.custTh(th.getColNames());
        } else if(result instanceof ElemCustTableRows){
            ElemCustTableRows rows = (ElemCustTableRows)result;
            bookmarkNode.bookmark.custTr(rows);
        } else if(result instanceof ElemTableRows){
            //处理表格，if the result for this bookmark is a Table Row Set
            ElemTableRows<?> rows = (ElemTableRows<?>)result;
            int size = rows.getRowCount();
            List<XWPFTableRow> templateRows = bookmarkNode.getTableTemplateRows();
            bookmarkNode.initTableTemplateRows();
            
            if(size==0){
                //clean table
                for(BookmarkTreeNode child:children){
                    child.remove();
                }
            }
            
            for(int i=size-1;i>=0;){
                if(templateRows==null){
                    break;
                }
                
                int sizeOfTemplateRows = templateRows.size();
                int iterSize = sizeOfTemplateRows<size?sizeOfTemplateRows:size;
                List<BookmarkTreeNode> notReplacedChildren = new ArrayList<BookmarkTreeNode>();
                notReplacedChildren.addAll(children);
                
                for(int j=iterSize-1;j>=0;j--){
                    if(i<0)break;
                    
                    XWPFTableRow templateRow = templateRows.get(j);
                    Object elem = rows.get(i);
                    //对应模板行每一行，迭代子节点
                    for(BookmarkTreeNode child:children){
                        if(child.bookmark.startTableRow==templateRow){
                            replaceBookmarks(elem,child);
                            notReplacedChildren.remove(child);
                        }
                    }
                    i--;
                }
                
                //对于从来没有被替换处理的模板行，我们要处理一下，把所有内容替换为空，否则对应内容会保留模板内容
                for(BookmarkTreeNode notReplaced:notReplacedChildren){
                    notReplaced.remove();
                }
                
                //滚动复制
                if(i>=0)bookmarkNode.copy();
            }                
        } else if(result instanceof ElemList){
            //if the result for this bookmark is a container (is a instance of ElemList), we need to deal with the children bookmarks of it
            ElemList<?> elemList = (ElemList<?>)result;
            int size = elemList.getSize();
            for(int i=0;i<size;i++){
                Object elem = elemList.get(i);
                //对应集合里面的每一行数据，迭代子节点
                for(BookmarkTreeNode child:children){
                    replaceBookmarks(elem,child);
                }
                
                //滚动复制
                if(i<size-1)bookmarkNode.copy();
            }
            
        }else{
            if(result == Elem.KEEP_ORIGINAL){
                //do nothing
            }else if(result instanceof ElemImage){
                //if it's image
                bookmarkNode.replaceImage(result.toString());
            }else if(result instanceof Double){
                bookmarkNode.replaceText(DocHelper.format(result));
            }else{
                //otherwise, the value for this bookmark is simple object, so we just replace the value in template.
                bookmarkNode.replaceText(result.toString());
            }
            
            //replace the children nodes of this bookmark
            //一般来说，如果一个bookmark节点有子节点的话，对应该节点的数据类型都是一个集合类型(ElemList),这里处理非集合类型的情况
            for(BookmarkTreeNode child:children){
                replaceBookmarks(result,child);
            }
        }
    }
    Object getProperty(Object bean, BookmarkTreeNode bookmarkNode) {
        String propPath = bookmarkNode.bookmark.propertyPath;
        if(bookmarkNode.parent!=null){
            String parentPropPath = bookmarkNode.parent.bookmark.propertyPath;
            if(!StringUtils.isBlank(parentPropPath)){
                //去掉父级路径
                propPath = propPath.replaceFirst("^"+parentPropPath+"_", ""); 
            }
        }
        return getProperty(bean,propPath);
    }
    Object getProperty(Object bean, String propPath){
        //绝对路径
        if(propPath.startsWith("R_")){
            bean = rootBean;
            propPath = propPath.replaceFirst("R_", "");
        }
        return getProperty(bean,propPath.split("_") );
        
    }
    Object getProperty(Object bean, String[] path) {
        if(bean==null)return null;
        if(path==null)return bean;
        if(path.length==0)return bean;
                   
        try {
            Object ret = null;
            if(StringUtils.isBlank(path[0])){
                ret = bean;
            }else{
                ret = PropertyUtils.getProperty(bean, path[0]);
            }               
            
            if(path.length==1){
                return ret;
            }
            path = Arrays.copyOfRange(path, 1, path.length);
            return getProperty(ret, path);
        } catch (Exception e) {
            //e.printStackTrace();
            System.err.println(e.getLocalizedMessage());
            return null;
        } 
    }

}