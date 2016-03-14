package cn.com.genechem.doc;


public class ElemTable<E> implements Elem {
    String name = null;
    int contrastGroup;
    ElemTableRows<E> rows;
    
    public ElemTable(String name){
        this.name = name;
        this.rows = new ElemTableRows<E>(); 
    }
    public String getName(){
        return name;
    }
 
    public int getContrastGroup() { return contrastGroup; }
    public void setContrastGroup(int contrastGroup) { this.contrastGroup = contrastGroup; }
    public ElemTableRows<E> getRows() {
        return rows;
    }

}
