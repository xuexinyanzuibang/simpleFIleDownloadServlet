package cn.com.genechem.doc;

public interface Elem {
    public static Elem KEEP_ORIGINAL=ElemOriginal.ORIGINAL;

}

class ElemOriginal implements Elem{
    static ElemOriginal ORIGINAL = new ElemOriginal();
    private ElemOriginal(){}
}
