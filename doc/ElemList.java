package cn.com.genechem.doc;

import java.util.ArrayList;

@SuppressWarnings("serial")
public class ElemList<E> extends ArrayList<E> implements Elem {
    public int getSize(){
        return size();
    }    
}
