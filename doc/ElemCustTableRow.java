package cn.com.genechem.doc;


public class ElemCustTableRow implements Elem {
    Object [] values;
    ElemCustTableRow(Object... values){
        this.values = values;
    }
    int getColCount(){
        return values.length;
    }
    Object getValue(int col){
        return values[col];
    }
}
