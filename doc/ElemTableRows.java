package cn.com.genechem.doc;

import java.util.ArrayList;

@SuppressWarnings("serial")
public class ElemTableRows<E> extends ArrayList<E> implements Elem {
    public int getRowCount(){
        return size();
    }    
}

