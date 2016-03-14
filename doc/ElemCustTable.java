package cn.com.genechem.doc;


public class ElemCustTable implements Elem {
    String name = null;
    ElemCustTableHead th;
    ElemCustTableRows rows;
    
    public ElemCustTable(String name){
        this.name = name;
        this.rows = new ElemCustTableRows();        
    }
    public String getName(){
        return name;
    }
    public void setTh(String ... cols){
        th = new ElemCustTableHead(cols);
    }
    public ElemCustTableHead getTh(){
        return th;
    }
    public int getColCount(){
        return th.getColCount();
    }
    public void addRow(Object... values){
        rows.add(new ElemCustTableRow(values));
    }
    public ElemCustTableRows getRows() {
        return rows;
    }
}
