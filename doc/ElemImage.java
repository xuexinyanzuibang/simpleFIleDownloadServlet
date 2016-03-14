package cn.com.genechem.doc;

public class ElemImage implements Elem {
    String src;
    public ElemImage(String src){
        this.src = src;
    }
    @Override
    public String toString(){
        return src;
    }
    
}
