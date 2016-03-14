package cn.com.genechem.doc;

import java.math.BigDecimal;

public class DocHelper {
    static String format(Object result){
        if(result instanceof Double){
            BigDecimal bd = new BigDecimal((Double)result);
            bd = bd.setScale(3,BigDecimal.ROUND_HALF_UP);
            return bd.toString();
        }
        
        return result+"";
    }
}
