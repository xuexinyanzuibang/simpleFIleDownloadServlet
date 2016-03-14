package cn.com.genechem.doc;

import java.util.ArrayList;

@SuppressWarnings("serial")
public class ElemCustTableRows extends ArrayList<ElemCustTableRow> implements Elem {
    public int getRowCount(){
        return size();
    }    
}

