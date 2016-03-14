package cn.com.genechem.doc;

public class ElemBreak implements Elem {
    public static final ElemBreak TEXT_WRAPPING  = new ElemBreak("\n"); 
    public static final ElemBreak PAGE_BREAK     = new ElemBreak("\r");
    
    private String txt="\n";
    private ElemBreak(String txt){
      this.txt=txt;  
    }
    
    @Override
    public String toString(){
        return txt;
    }
}
