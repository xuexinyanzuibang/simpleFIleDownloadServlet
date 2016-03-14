package cn.com.genechem.doc;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

import cn.com.genechem.doc.DocTemplate.Bookmark;

public class DocTemplate extends XWPFDocument {
    static boolean           debug;
    static{
        //check if enabled the debug mode
        debug = "true".equals(System.getProperty("DocTemplate.debug"));
        //debug=true;
        Log.debug("!!!! DocTemplate Debug Mode !!!!");
    }
    
	BookmarksMgr             bookmarksMgr = new BookmarksMgr();
	Set<IBodyElement>        removedElems = new HashSet<IBodyElement>();
    
	DocTemplate(InputStream in) throws IOException{
		super(in);
		loadBoomarks();
	}
    List<BookmarkTreeNode> getBookmarkTreeNodes(){
    	return bookmarksMgr.getBookmarkTreeNodes();
    }
    BookmarkTreeNode getBookmarkRootNode(){
        return bookmarksMgr.bookmarkRootNode;
    }
    private void loadBoomarks() {
        loadTablesBookmarks(getTables());
        loadParasBookmarks(getParagraphs());
        loadBodyBookmarks();
        
        bookmarksMgr.buildBookmarkTree();
    }
    private void loadBodyBookmarks() {
    	 CTBody body = getDocument().getBody();
         loadBoomarks(body.getBookmarkStartList(), body.getBookmarkEndList(),null,null,null);
    }
    private void loadTablesBookmarks(List<XWPFTable> tableList) {
        for (XWPFTable table : tableList) {
            List<XWPFTableRow> rowList = table.getRows();
            for (XWPFTableRow row : rowList) {
                List<XWPFTableCell> cellList = row.getTableCells();
                for (XWPFTableCell cell : cellList) {
                    loadParasBookmarks(cell.getParagraphs(),table,row);
                    loadBoomarks(cell.getCTTc().getBookmarkStartList(), cell.getCTTc().getBookmarkEndList(),null,table,row);
                }
                loadBoomarks(row.getCtRow().getBookmarkStartList(), row.getCtRow().getBookmarkEndList(),null,table,null);
            }
            loadBoomarks(table.getCTTbl().getBookmarkStartList(), table.getCTTbl().getBookmarkEndList(),null,table,null);
        }
    }

    private void loadParasBookmarks(List<XWPFParagraph> paraList) {
        loadParasBookmarks(paraList,null,null);   
    }
    private void loadParasBookmarks(List<XWPFParagraph> paraList,XWPFTable table,XWPFTableRow tableRow) {

        for (XWPFParagraph para : paraList) {
            List<CTBookmark> bookmarkStartList = para.getCTP().getBookmarkStartList();
            List<CTMarkupRange> bookmarkEndList = para.getCTP().getBookmarkEndList();
            loadBoomarks(bookmarkStartList, bookmarkEndList,para,table,tableRow);
        }
    }

    private void loadBoomarks(List<CTBookmark> bookmarkStartList,List<CTMarkupRange> bookmarkEndList,XWPFParagraph para,XWPFTable table,XWPFTableRow tableRow) {
        for (CTBookmark mark : bookmarkStartList) {
        	Log.debug("Start " + mark.getId() + " " + mark.getName());
            bookmarksMgr.bookmarkStart(mark,para,table,tableRow);
        }

        for (CTMarkupRange mark : bookmarkEndList) {
            Log.debug("End " + mark.getId());
            // ctMarkup.getDomNode().getNodeName()
            // ctMarkup.dump();

            bookmarksMgr.bookmarkEnd(mark,para,table,tableRow);
            // cursor.selectPath("./*");
            // Log.debug(ctMarkup);
        }

    }
    
    void cleanRemovedElems(){
    	for(IBodyElement e:removedElems){
    		int pos = this.getBodyElements().indexOf(e);
    		this.removeBodyElement(pos);
    	}
    	removedElems.clear();
    }
    
    class BookmarksMgr {
        Map<BigInteger, Bookmark>   bookmarksById    	= new HashMap<BigInteger, Bookmark>();
        List<BookmarkCursor>        bookmarkCursors  	= new ArrayList<BookmarkCursor>();
        BookmarkTreeNode            bookmarkRootNode 	= new BookmarkTreeNode(new Bookmark(BigInteger.valueOf(-1),""));
        BookmarkTreeBuilder 		bookmarkTreeBuilder = new BookmarkTreeBuilder(bookmarkRootNode);
        
        List<BookmarkTreeNode> getBookmarkTreeNodes(){
            return bookmarkRootNode.children;
        }
        void bookmarkStart(CTBookmark start,XWPFParagraph para,XWPFTable table,XWPFTableRow tableRow) {
            Bookmark bookmark = registerBookmark(start.getId()).start(start, para,table,tableRow);
            addBookmarkCursor(bookmark,bookmark.start.newCursor(),BookmarkCursor.Type.START);
        }
        
        void bookmarkEnd(CTMarkupRange end, XWPFParagraph para,XWPFTable table,XWPFTableRow tableRow) {
            Bookmark bookmark = registerBookmark(end.getId()).end(end, para,table,tableRow);
            addBookmarkCursor(bookmark,bookmark.end.newCursor(),BookmarkCursor.Type.END);
        }
        //void addBook(Bookmark bookmark) {
        Bookmark getBookmarkById(BigInteger id) {
            return bookmarksById.get(id);
        }
        
        void buildBookmarkTree(){
            Collections.sort(bookmarkCursors, new BookmarkCursorComparator());
            
            for(BookmarkCursor cursor:bookmarkCursors){
                if("_GoBack".equals(cursor.bookmark.name)){
                    //忽略word内置书签
                    continue;
                }
                if(cursor.bookmark.name.startsWith("_Toc")){
                    //忽略word内置书签
                    continue;
                }
                Log.debug(cursor.bookmark.id+", "+cursor.bookmark.name+", "+cursor.type);
                bookmarkTreeBuilder.procCursor(cursor);
                cursor.dispose();
            }
            bookmarkCursors.clear();
            bookmarkCursors = null;
            Log.debug(bookmarkRootNode.dump());
            
        }
        
        private Bookmark registerBookmark(BigInteger id) {
            Bookmark bookmark = getBookmarkById(id);
            if (null == bookmark) {
                bookmark = new Bookmark(id);
            }
            bookmarksById.put(bookmark.id, bookmark);
            return bookmark;
        }
        private void addBookmarkCursor(Bookmark bookmark, XmlCursor xmlCursor,BookmarkCursor.Type type){
            bookmarkCursors.add(new BookmarkCursor(bookmark, xmlCursor,type));
        }
        
    }
    
    class Bookmark {
        BigInteger                  id;
        String                      name;
        String                      propertyPath;
        String                      propertySuffix;
        CTBookmark                  start;
        CTMarkupRange               end;
        IBodyElement                startElem;
        IBodyElement                endElem;
        List<IBodyElement>          elemList;
        List<XWPFRun>               runList;
        XWPFRun                     run;
        XWPFTable                   table;
		XWPFTableRow                startTableRow;
		XWPFTableRow                endTableRowEdge;
        List<XWPFTableRow>          tableTemplateRows;
        int                          cx;
        int                          cy;
        
        Bookmark(BigInteger id) {
            this.id = id;
        }
        Bookmark(BigInteger id,String name) {
            this.id = id;
            this.name = name;
            parseBookmarkName();
        }
        Bookmark start(CTBookmark start,XWPFParagraph para,XWPFTable table,XWPFTableRow tableRow) {
            this.start = start;
            this.name = this.start.getName();
            parseBookmarkName();
            this.startElem=para;
            this.table = table;
            this.startTableRow = tableRow;
            if(null==this.startElem){
                //this bookmark is started from body, so we find the start para below the bookmark tag, just go follow this
                //find the start para, start table
                XmlCursor cursor = this.start.newCursor();
                
                //max 10 tries
                CTP ctp = null; // CTP for para
                CTTbl ctTbl = null; // CTTbl for table
                for(int i=1;i<10;i++){
                    if(!cursor.toNextSibling()){
                        break;
                    }
                    XmlObject xmlObj = cursor.getObject();
                    if(xmlObj instanceof CTP){
                        ctp = (CTP)xmlObj;
                        break;
                    }else if(xmlObj instanceof CTTbl){
                        ctTbl = (CTTbl)ctTbl;
                    }
                }
                
                if(ctp!=null){
                    this.startElem = getParagraph(ctp);
                }
                if(ctTbl!=null){
                    this.startElem = getTable(ctTbl);
                }
                
                cursor.dispose();
            }
            return this;
        }
        private void parseBookmarkName(){
            int idx = name.indexOf("__");
            //remove the __1,__2,__x from the bookmark name
            if(-1 != idx){
                propertyPath = name.substring(0,idx);
                propertySuffix = name.substring(idx+2);
            }else{
                propertyPath = name;
            }
        }
        Bookmark end(CTMarkupRange end, XWPFParagraph para,XWPFTable table,XWPFTableRow tableRow) {
            this.end=end;
            if(this.table==null){
                this.table = table;
            }
            if(null==para){
                //find the end para
                XmlCursor cursor = this.end.newCursor();
                
                //max 10 tries
                CTP ctp = null;
                CTTbl ctTbl = null; // CTTbl for table
                CTRow ctRow = null; // CTTbl for table row
                for(int i=1;i<10;i++){
                    if(!cursor.toPrevSibling()){
                        break;
                    }
                    XmlObject xmlObj = cursor.getObject();
                    if(xmlObj instanceof CTP){
                        ctp = (CTP)xmlObj;
                        break;
                    }
                    else if(xmlObj instanceof CTTbl){
                        ctTbl = (CTTbl)xmlObj;
                        break;
                    }
                    else if(xmlObj instanceof CTRow){
                        ctRow = (CTRow)xmlObj;
                        break;
                    }

                }
                
                if(ctp!=null){
                    this.endElem = getParagraph(ctp);
                }
                if(ctTbl!=null){
                    this.endElem = getTable(ctTbl);
                }
                if(ctRow!=null&&table!=null){
                    this.endTableRowEdge = table.getRow(ctRow);
                    //fix bug
                    if(null!=this.endTableRowEdge){
                        List<XWPFTableRow> tableRows = table.getRows();
                        int endTableRowIdx = tableRows.indexOf(endTableRowEdge);
                        if(endTableRowIdx<tableRows.size()-1){
                            this.endTableRowEdge=tableRows.get(endTableRowIdx+1);
                        }else{
                            this.endTableRowEdge=null;
                        }
                    }
                }

                cursor.dispose();
                
            }else{
                this.endElem=para;
            }
            
            if(this.startElem instanceof XWPFParagraph && this.startElem==this.endElem){
                this.endElem=null;
                this.runList = new ArrayList<XWPFRun>();
                //this bookmark is in a para, so we find the Runs for it
                XmlCursor cursor = this.start.newCursor();
                XmlCursor endCursor = this.end.newCursor();
                
                while(cursor.toNextSibling()&&!cursor.isAtSamePositionAs(endCursor)){
                    XmlObject xmlObj = cursor.getObject();
                    if(xmlObj instanceof CTR){
                        CTR r = (CTR)xmlObj;
                        XWPFRun run = ((XWPFParagraph)this.startElem).getRun(r);
                        if(run==null){
                            continue;
                        }
                        runList.add(run);
                    }                    
                }
                if(runList.size()==0){
                	runList = null;
                }
                cursor.dispose();
                endCursor.dispose();
      
            }else{
                List<IBodyElement> bodyElems = getBodyElements();
                int startIdx = bodyElems.indexOf(startElem);
                int endIdx = bodyElems.indexOf(endElem);
                if (startIdx != -1 && endIdx != -1) {
                    elemList = new ArrayList<IBodyElement>();
                    Log.debug(this.toString() + startIdx + "," + endIdx);
                    for (int i = startIdx; i <= endIdx; i++) {
                        elemList.add(bodyElems.get(i));
                    }
                }
                
                //for table cell
                if(this.endElem instanceof XWPFParagraph && this.startElem!=this.endElem &&this.startTableRow!=null&& tableRow!=null){
                    List<XWPFTableRow> rows = table.getRows();
                    int startRowIdx = rows.indexOf(startTableRow);
                    int endRowIdx = rows.indexOf(tableRow);
                    if(startRowIdx<endRowIdx){
                        this.endTableRowEdge=rows.get(endRowIdx);
                    }
                }
            }
            return this;
        }

        void copy(){
            copyTableRows();
            copyBodyElems();
        }
        private void copyBodyElems(){
            if(elemList==null){
                return;
            }
            int size = elemList.size();
            IBodyElement newElem = startElem;
            //XWPFParagraph newP = startElem;
            for(int i=size-1;i>=0;i--){
                IBodyElement e = elemList.get(i);
                if(removedElems.contains(e)){
                	continue;
                }
                XmlCursor cursor = null;
                if(newElem instanceof XWPFParagraph){
                    cursor = ((XWPFParagraph) newElem).getCTP().newCursor();
                }else if(newElem instanceof XWPFTable){
                    cursor = ((XWPFTable) newElem).getCTTbl().newCursor(); 
                }
                
                if(e instanceof XWPFParagraph){
                    newElem = insertNewParagraph(cursor);
                    cursor.dispose();
                    getDocument().getBody().setPArray(paragraphs.indexOf(newElem), (CTP) ((XWPFParagraph)e).getCTP().copy());
                }else if(e instanceof XWPFTable){
                    newElem = insertNewTbl(cursor);
                    cursor.dispose();
                    getDocument().getBody().setTblArray(tables.indexOf(newElem), (CTTbl) ((XWPFTable)e).getCTTbl().copy());
                }
                
                //setParagraph(paragraph, pos);
                removeBookmarks(newElem);
            }
        }
        @SuppressWarnings("unused")
        private void copyTableRows(){
            if(tableTemplateRows==null){
                return;
            }
            if(null==table){
            	return;
            }
            if(removedElems.contains(table)){
            	return;
            }
            int size = tableTemplateRows.size();
            int insertPos = table.getRows().indexOf(tableTemplateRows.get(size-1))+1;
            for(int i=size-1;i>=0;i--){
                XWPFTableRow newRow = table.insertNewTableRow(insertPos);
                //XWPFTableRow newRow = table.insertNewTableRow(1);
                table.getCTTbl().setTrArray(insertPos, (CTRow)tableTemplateRows.get(i).getCtRow().copy());

            }
        }
        private void removeBodyElems(){
        	/*

            if(elemList==null){
                return;
            }
            int size = elemList.size();

            for(int i=size-1;i>=0;i--){
                IBodyElement e = elemList.get(i);
                int pos = getBodyElements().indexOf(e);
                removeBodyElement(pos);
            }
            elemList=null;
            table=null;
            tableTemplateRows=null;
            runList=null;
            */
            if(elemList==null){
                return;
            }
            removedElems.addAll(elemList);
        }
        private void resetBodyElems(){

            if(elemList==null){
                return;
            }
            removedElems.removeAll(elemList);
        }
        private void removeTable(){
        	/*
        	if(checkTable()==null){
        		return;
        	}
        	runList = null;
            if(tableTemplateRows==null){
                return;
            }
            int tablePos = getBodyElements().indexOf(table);
            removeBodyElement(tablePos);
            */
        	if(checkTable()==null){
        		return;
        	}
            if(tableTemplateRows==null){
                return;
            }
            removedElems.add(table);
        }
        private void resetTable(){

        	if(checkTable()==null){
        		return;
        	}
            if(tableTemplateRows==null){
                return;
            }
            removedElems.remove(table);
        }
        void reset(){
        	resetTable();
        	resetBodyElems();
        }
        
        XWPFTable checkTable() {
        	if(table==null)return null;
        	//check if this table is removed
        	int tablePos = getBodyElements().indexOf(table);
        	if(-1==tablePos){
        		table = null;
        	}
			return table;
		}

        
        XWPFRun findImageRun(){
            /*
            if(getBodyElements().indexOf(startElem)==-1){
            	// this part has been removed
            	return null;
            }*/
            
            XWPFRun run = null;
            //try to find the template image in the book mark
            List<XWPFRun> runs = runList;
            if(runs==null){
                return null;
            }
            
            
            CTR ctr = null;
            CTDrawing drawing = null;
            int size = -1;
            for(XWPFRun r:runs){
                ctr = r.getCTR();
                List<CTDrawing> drawings = ctr.getDrawingList();
                size=drawings.size();
                if(size!=0){
                    drawing = drawings.get(0);
                    run = r;
                    break;
                }
            }
            
            if(run==null){
                return null;
            }
            
            List<CTInline> inlines = drawing.getInlineList();
            CTInline inline = inlines.get(0);
            CTPositiveSize2D posSize = inline.getExtent();
            //CTEffectExtent effctExt = inline.getEffectExtent();
            
            cx = (int)posSize.getCx();
            cy = (int)posSize.getCy();
            /*
            long effctExtB = effctExt.getB();
            long effctExtL = effctExt.getL();
            long effctExtR = effctExt.getR();
            long effctExtT = effctExt.getT();
            */
            
            return run;
        }
        void replaceImage(String image){
            if(run==null){
            	run = findImageRun();
            }
            
            if(run==null){
                //throw new RuntimeException("No drawing found in template doc.");
                //Log.debug("No drawing found for bookmark "+this.name);
            	return;
            }
            
            //remove the old image
            CTR ctr = run.getCTR();
            List<CTDrawing> drawings = ctr.getDrawingList();
            int size=drawings.size();
            for(int i=0;i<size;i++){
                ctr.removeDrawing(i);
            }

            //replace with the new image
            if(StringUtils.isBlank(image)){
                run.setText(" ", 0);
                return;
            }
            File file = new File(image);;
            try(InputStream is = new FileInputStream(file)){
                run.addPicture(is, getImageType(image), file.getName(), cx, cy);
                /*
                CTDrawing newDrawing = ctr.getDrawingList().get(0);
                List<CTInline> newInlines = newDrawing.getInlineList();
                CTInline newInline = newInlines.get(0);
                newInline.addNewEffectExtent().setB(effctExtB);
                newInline.addNewEffectExtent().setL(effctExtL);
                newInline.addNewEffectExtent().setR(effctExtR);
                newInline.addNewEffectExtent().setT(effctExtT);
                */
                
                //DocTemplate.this.removeRelation(part, removeUnusedParts)
            } catch (Exception e) {
               throw new RuntimeException(e);
            }
        }
        private void clearText(XWPFRun run){
            List<CTText> texts = run.getCTR().getTList();
            int size = texts.size();
            for(int i=size-1;i>=0;i--){
                run.getCTR().removeT(i);
            }
            removeBr(run);
        }
        private void removeBr(XWPFRun run){
            List<CTBr> brs = run.getCTR().getBrList();
            int size = brs.size();
            for(int i=size-1;i>=0;i--){
                run.getCTR().removeBr(i);
            }
        }
        void replaceText(String value){
            List<XWPFRun> runs = runList;
            List<XWPFRun> removedRuns = new ArrayList<XWPFRun>();
            if(runs==null){
                return;
            }

            int size = runs.size();
            boolean replaced = false;
            
            for(int i=0;i<size;i++){
                XWPFRun run = runs.get(i);
                if(replaced){
                    XWPFParagraph para = ((XWPFParagraph)this.startElem);
                    List<XWPFRun> rs = para.getRuns();
                    int idx = rs.indexOf(run);
                    if(idx!=-1){
                        replaceText(run,null);
                        para.removeRun(idx);
                        removedRuns.add(run);
                    }
                    
                    continue;
                }
                
                if(i==size-1){
                    //last one
                    replaceText(run,value);
                }else{
                	if(!StringUtils.isBlank(run.getText(0))){
                		replaceText(run,value);
                		replaced = true;
                	}
                }
            }
            
            runs.removeAll(removedRuns);
        }
        private void replaceText(XWPFRun run, String txt){
            if(null==txt){
                txt="";
            }
            clearText(run);
            
            if("\r".equals(txt)){
                //add page break;
                run.addBreak(BreakType.PAGE);
                return;
            }
            
            String[] txts = txt.split("\n");
            for(int i=0;i<txts.length;i++){
                String value = txts[i];
                run.setText(value, i);
                if(i<txts.length-1){
                    run.addBreak();
                }
            }
        }

        void remove(){
            //标记已经删除的yuansu
            removeBodyElems();
            //标记已经删除的表
            removeTable();
            //清空文本
            replaceText(null);
            //清空图像
            replaceImage(null);
        }
        private void initCustTable(){
            
            List<XWPFTableRow> rowList = table.getRows();
            int size = rowList.size();
            for(int i=size-1;i>=0;i--){
                XWPFTableRow row = rowList.get(i);
                if(row==startTableRow){
                    return;
                }
                table.removeRow(i);
            }
        }
        private void copyCustTr(){
            int insertPos =table.getRows().size();
            XWPFTableRow newRow = table.insertNewTableRow(table.getRows().size());
            table.getCTTbl().setTrArray(insertPos, (CTRow)startTableRow.getCtRow().copy());  
        }
        private void custTr(ElemCustTableRow row){
            int colCount = row.getColCount();
            for(int j=0;j<colCount;j++){
                if(j==0){
                    //nothing
                }else{
                    replaceText(DocHelper.format(row.getValue(j)));
                    startTableRow.getCtRow().setTcArray(j, (CTTc)startTableRow.getCtRow().getTcArray(0).copy());
                    
                }
                
                if(j==colCount-1){
                	replaceText(DocHelper.format(row.getValue(0)));
                }
            }
        }
        void custTr(ElemCustTableRows rows){
            //Log.debug(rows.getRowCount());
            initCustTable();
            int templateRow = table.getRows().indexOf(startTableRow);
            if(-1==templateRow)return;
            
            int rowCount = rows.getRowCount();
            for(int i=0;i<rowCount;i++){
                if(i==0){
                    //nothing
                }else{
                    custTr(rows.get(i));
                    copyCustTr();
                }
                
                if(i==rowCount-1){
                	custTr(rows.get(0));
                }
                
            }
        }
        void custTh(String[] cols){
            
            for(int i=0;i<cols.length;i++){
                if(i==0){
                    //replaceText(cols[0]);
                }else{
                    replaceText(cols[i]);
                    this.table.addNewCol();
                    startTableRow.getCtRow().setTcArray(i, (CTTc)startTableRow.getCtRow().getTcArray(0).copy());
                    //this.table.getCTTbl().get
                    //table.getCTTbl().setTrArray(insertPos, (CTRow)tableTemplateRows.get(i).getCtRow().copy());
                    //table.getCTTbl().sett
                    //table.getCTTbl().setTblGrid(tblGrid);
                    if(i==cols.length-1){
                        replaceText(cols[0]);
                    }
                }
            }
        }
        @SuppressWarnings("unused")
        private void removeBookmarks(IBodyElement elem){
            if(elem instanceof XWPFParagraph){
                CTP ctp = ((XWPFParagraph)elem).getCTP();
                while(ctp.getBookmarkStartList().size()!=0){
                    ctp.removeBookmarkStart(0);
                }
                while(ctp.getBookmarkEndList().size()!=0){
                    ctp.removeBookmarkEnd(0);
                }
            }
            
            if(elem instanceof XWPFTable){
                XWPFTable tbl = (XWPFTable)elem;
                
                List<XWPFTableRow> rowList = tbl.getRows();
                for (XWPFTableRow row : rowList) {
                    List<XWPFTableCell> cellList = row.getTableCells();
                    for (XWPFTableCell cell : cellList) {
                        /*
                        List<XWPFParagraph> paras = cell.getParagraphs();
                        for(XWPFParagraph p:paras){
                            removeBookmarks(p);
                        }
                        
                        
                        CTTc ctTc = cell.getCTTc();
                        while(ctTc.getBookmarkStartList().size()!=0){
                            ctTc.removeBookmarkStart(0);
                        }
                        while(ctTc.getBookmarkEndList().size()!=0){
                            ctTc.removeBookmarkEnd(0);
                        }*/
                    }
                    
                    //remove the row level bookmarks for the copy row
                    /*
                    CTRow ctRow = row.getCtRow();
                    while(ctRow.getBookmarkStartList().size()!=0){
                        ctRow.removeBookmarkStart(0);
                    }
                    while(ctRow.getBookmarkEndList().size()!=0){
                        ctRow.removeBookmarkEnd(0);
                    }*/
                }
                
                //Remove the table level bookmarks for the copy table
                CTTbl ctTbl = tbl.getCTTbl();
                while(ctTbl.getBookmarkStartList().size()!=0){
                    ctTbl.removeBookmarkStart(0);
                }
                while(ctTbl.getBookmarkEndList().size()!=0){
                    ctTbl.removeBookmarkEnd(0);
                }
                
            }
            
        }
        @Override
        public String toString(){
            String ret = "["+id+","+name+",";
            if(this.startElem!=null){
                //ret += getPosOfParagraph(startElem);
                ret += ".";
                ret += paragraphs.indexOf(startElem);
            }
            if(this.endElem!=null){
                ret += "~";
                //ret += getPosOfParagraph(endElem);
                ret += ".";
                ret += paragraphs.indexOf(endElem);
            }

            ret += "]";
            return ret;
        }
    }
    private int getImageType(String imgFile){
        int format;

        if(imgFile.endsWith(".emf")) format = XWPFDocument.PICTURE_TYPE_EMF;
        else if(imgFile.endsWith(".wmf")) format = XWPFDocument.PICTURE_TYPE_WMF;
        else if(imgFile.endsWith(".pict")) format = XWPFDocument.PICTURE_TYPE_PICT;
        else if(imgFile.endsWith(".jpeg") || imgFile.endsWith(".jpg")) format = XWPFDocument.PICTURE_TYPE_JPEG;
        else if(imgFile.endsWith(".png")) format = XWPFDocument.PICTURE_TYPE_PNG;
        else if(imgFile.endsWith(".dib")) format = XWPFDocument.PICTURE_TYPE_DIB;
        else if(imgFile.endsWith(".gif")) format = XWPFDocument.PICTURE_TYPE_GIF;
        else if(imgFile.endsWith(".tiff")) format = XWPFDocument.PICTURE_TYPE_TIFF;
        else if(imgFile.endsWith(".eps")) format = XWPFDocument.PICTURE_TYPE_EPS;
        else if(imgFile.endsWith(".bmp")) format = XWPFDocument.PICTURE_TYPE_BMP;
        else if(imgFile.endsWith(".wpg")) format = XWPFDocument.PICTURE_TYPE_WPG;
        else {
            throw new RuntimeException("Unsupported picture: " + imgFile +
                    ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");
        }
        return format;
    }
}



class BookmarkCursor implements Comparable<BookmarkCursor>{
    enum Type {START,END}
    Type type;
    Bookmark bookmark;
    XmlCursor xmlCursor;
    
    BookmarkCursor(Bookmark bookmark, XmlCursor xmlCursor,Type type){
        this.bookmark = bookmark;
        this.xmlCursor = xmlCursor;
        this.type = type;
    }
    void dispose(){
    	xmlCursor.dispose();
    }
    @Override
    public int compareTo(BookmarkCursor o) {
        String thisName = this.bookmark.name;
        String otherName = o.bookmark.name;
        //名字相同的，end 总大于start
        if(thisName.equals(otherName)){
            return this.type.ordinal()-o.type.ordinal();
        }
        
        if(this.type==Type.START && this.type==o.type){
            if(thisName.startsWith(otherName+"_")){
                return 1;
            }
            if(otherName.startsWith(thisName+"_")){
                return -1;
            }
        }
        if(this.type==Type.END && this.type==o.type){
            if(thisName.startsWith(otherName+"_")){
                return -1;
            }
            if(otherName.startsWith(thisName+"_")){
                return 1;
            }
        }
        
        return this.xmlCursor.comparePosition(o.xmlCursor);
    }
    
}
class BookmarkCursorComparator implements Comparator<BookmarkCursor>{
    @Override
    public int compare(BookmarkCursor o1, BookmarkCursor o2) {
        return o1.compareTo(o2);
    }
    
}
class BookmarkTreeBuilder{
    Map<BigInteger, BookmarkTreeNode>   bookmarkNodesById       = new HashMap<BigInteger, BookmarkTreeNode>();
    Map<String, BookmarkTreeNode>  		bookmarkNodesByName		= new HashMap<String, BookmarkTreeNode>();
    BookmarkTreeNode                    currentNode 			= null;
    
    BookmarkTreeBuilder(BookmarkTreeNode currentNode){
        this.currentNode = currentNode;
        registerNode(currentNode);
    }
    void procCursor(BookmarkCursor cursor){
        if(cursor.type==BookmarkCursor.Type.START){
            BookmarkTreeNode node = new BookmarkTreeNode(currentNode, cursor.bookmark);
            registerNode(node);
            currentNode = node;
        }else{
            if(cursor.bookmark.id.equals(currentNode.bookmark.id)){
                currentNode.complete();
                currentNode = currentNode.parent;
            }else{
                BookmarkTreeNode node = bookmarkNodesById.get(cursor.bookmark.id);
                node.complete();
            }
        }
    }
    void registerNode(BookmarkTreeNode currentNode){
        bookmarkNodesById.put(currentNode.bookmark.id, currentNode);
        bookmarkNodesByName.put(currentNode.bookmark.name, currentNode);
    }
}

class BookmarkTreeNode{
    boolean isComplete;
    Bookmark bookmark;
    BookmarkTreeNode parent;
    List<BookmarkTreeNode> children = new ArrayList<BookmarkTreeNode>();

    BookmarkTreeNode(Bookmark data){
        this.bookmark = data;
    }
    BookmarkTreeNode(BookmarkTreeNode parent,Bookmark data){
        this.bookmark = data;
        parent.appendChild(this);
    }

    void appendChild(BookmarkTreeNode child){
        children.add(child);
        child.parent = this;
    }
    
    void complete(){
        isComplete = true;
        
        //move all un-complete bookmark to its parent
        List<BookmarkTreeNode> removed = new ArrayList<BookmarkTreeNode>(); 
        for(BookmarkTreeNode child:children){
            if(child.isComplete){
                continue;
            }
            removed.add(child);
            parent.appendChild(child);
        }
        
        children.removeAll(removed);
    }
    void reset(){
        bookmark.reset();
    }
    void copy(){
        bookmark.copy();
    }
    void replaceText(String value){
        bookmark.replaceText(value);
    }
    void replaceImage(String image){
        bookmark.replaceImage(image);
    }
    void remove(){
    	initTableTemplateRows();//if this bookmark is for a table, we should init table firstly
        bookmark.remove();
    }
    void initTableTemplateRows(){
        List<XWPFTableRow> templateRows = getTableTemplateRows();
        if(templateRows==null)return;
        if(bookmark.checkTable()==null)return;
        
        List<XWPFTableRow> rowList = bookmark.table.getRows();
        int size = rowList.size();
        int start = size-1;
        
        if(bookmark.endTableRowEdge!=null){
            int idx = rowList.indexOf(bookmark.endTableRowEdge);
            if(-1!=idx){
                start = idx-1;
            }
        }
        
        for(int i=start;i>=0;i--){
            XWPFTableRow row = rowList.get(i);
            if(templateRows.contains(row)){
                break;
            }
            bookmark.table.removeRow(i);
        }
    }
    List<XWPFTableRow> getTableTemplateRows(){
        if(bookmark.tableTemplateRows == null){
            bookmark.tableTemplateRows = new ArrayList<XWPFTableRow>();
            //List<XWPFTableRow> rows = bookmark.table.getRows();
            for(BookmarkTreeNode child:children){
                //int idx = rows.indexOf(child.bookmark.tableRow);
            	if(child.bookmark.startTableRow==null){
            		continue;
            	}
                if(!bookmark.tableTemplateRows.contains(child.bookmark.startTableRow)){
                    bookmark.tableTemplateRows.add(child.bookmark.startTableRow);
                }
            }
            
            if(bookmark.tableTemplateRows.size()==0){
                bookmark.tableTemplateRows=null;
            }
        }
        
        return  bookmark.tableTemplateRows;
    }
    /*
    void fetchTableList(List<XWPFTable> list){
        if(bookmark.table!=null){
            if(!list.contains(bookmark.table)){
                list.add(bookmark.table);
            }
            for(BookmarkTreeNode child:children){
                child.fetchTableList(list); 
            }
        }
    }*/
    
    String dump(){
        return dump(0);
    }
    String dump(int level){
        String str = "";
        for(int i=0;i<level;i++){
            str  += "    ";
        }
        String ret = ""+str+""+this.toString();
        if(children.size()>0){
        	ret += "\n"+str+"{";
        	for(BookmarkTreeNode child:children){
        		ret += "\n"+child.dump(level+1);
        	}
        	ret += "\n"+str+"}";
        }
        return ret;
    }
    @Override
    public String toString(){
        return bookmark.toString();
    }
}

class Log{
    static public void debug(String msg){
        if(DocTemplate.debug){
            System.out.println(msg);
        }
    }
}

