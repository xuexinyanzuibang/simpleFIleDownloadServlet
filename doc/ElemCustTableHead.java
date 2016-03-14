package cn.com.genechem.doc;


public class ElemCustTableHead implements Elem {
    String [] colNames;
    ElemCustTableHead(String... colNames){
        this.colNames=colNames;
    }
    int getColCount(){
        return colNames.length;
    }
    public String[] getColNames() {
        return colNames;
    }
}
