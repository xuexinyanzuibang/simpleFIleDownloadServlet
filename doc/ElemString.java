package cn.com.genechem.doc;

public class ElemString implements Elem {
    private String str;
    public ElemString(String str){
        this.str=str;
    }
    @Override
    public String toString(){
        return str;
    }
}
