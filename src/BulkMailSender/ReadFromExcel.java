package BulkMailSender;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ReadFromExcel {
    POIFSFileSystem fs;
    OutputStream os;
	HSSFWorkbook wb;
	HSSFSheet sheet;
    HSSFRow row;
    HSSFCell cell;
	String[] emailCCValue;
	String emailBCCValue;
	String fileAttachValue;
	String excelEmailAttachFilesExt;
	String excelEmailAttachFilesFolder;
	String excelFilePath;

	int rows; // No of rows
    int rowIdx; // current row
    int cols = 0; // No of columns
    int emailColIdx = -1, emailBCCColIdx = -1, fileAttachColIdx = -1, emailCCNo = -1;
    int[] emailCCColIdx = new int[10];
    
	public ReadFromExcel(String excelFilePath, int excelheaderRow, String wbName, int startFrom, String emailFieldName, 
						 String emailCCFieldName, String emailBCCFieldName, String filenameFieldName,
						 String excelEmailAttachFilesFolder, String excelEmailAttachFilesExt) throws IOException
	{
		this.excelFilePath = excelFilePath;
	    fs = new POIFSFileSystem(new FileInputStream(excelFilePath));
	    wb = new HSSFWorkbook(fs);
	    sheet = wb.getSheet(wbName);

	    rows = sheet.getPhysicalNumberOfRows();
	    String[] emailCCList = emailCCFieldName.split(";");

	    this.excelEmailAttachFilesFolder = excelEmailAttachFilesFolder;
	    this.excelEmailAttachFilesExt = excelEmailAttachFilesExt;
	    
    	row = sheet.getRow(excelheaderRow);
	    for(int col = 0; (col < row.getLastCellNum()); col++)
	    {
	    	cell = row.getCell(col);
	    	if (cell == null)
	    		break;
	    	if (cell.getStringCellValue().compareTo(emailFieldName) == 0)
	    	{
	    		emailColIdx = col;
	    	}
	    	else if ((cell.getStringCellValue().compareTo(emailBCCFieldName) == 0) && (emailBCCFieldName.compareTo("") != 0))
	    	{
	    		emailBCCColIdx = col;
	    	}
	    	else if ((cell.getStringCellValue().compareTo(filenameFieldName) == 0) && (filenameFieldName.compareTo("") != 0))
	    	{
	    		fileAttachColIdx = col;
	    	}
	    	else
	    	{
	    		for(int y = 0; y < emailCCList.length; y++)
	    		{
		    		if ((cell.getStringCellValue().compareTo(emailCCList[y]) == 0) && (emailCCFieldName.compareTo("") != 0))
		    		{
		    			emailCCColIdx[++emailCCNo] = col;
		    			break;
		    		}
	    		}
	    	}

	    }
	    if (emailCCNo == -1)
	    {
	    	emailCCValue = new String[1];
	    	emailCCValue[0] = emailCCFieldName;
	    }
	    else
	    {
	    	emailCCValue = new String[++emailCCNo];
	    }
	    if (emailBCCColIdx == -1)
	    	emailBCCValue = emailBCCFieldName;
	    if (fileAttachColIdx == -1)
	    	fileAttachValue = filenameFieldName;

	    rowIdx = startFrom;
	}
	
	public HSSFRow getNextRow()
	{
	    if (rowIdx <= rows)
	    {
	    	row = sheet.getRow(rowIdx++);
	    }
	    else
	    {
	    	row = null;
	    }
	    return row;
	}
	
	public String getEmail()
	{
		return getField(emailColIdx).replaceAll(";", ",");
	}

	public String getEmailCCValue() 
	{
    	String retVal = "";
	    if (emailCCNo == -1)
			retVal = emailCCValue[0];
	    else
	    {
	    	String sep = "";
	    	for(int i = 0; i < emailCCNo; i++)
	    	{
	    		if ((getField(emailCCColIdx[i]) == null) || (getField(emailCCColIdx[i]).trim().compareTo("") == 0))
	    			continue;
	    		retVal += sep + getField(emailCCColIdx[i]);
	    		sep = ",";
	    	}
	    	retVal = retVal.replaceAll(";", ",");
	    }
	    return retVal;
	}

	public String getEmailBCCValue() {
	    if (emailBCCColIdx == -1)
	    	return emailBCCValue;
	    else
	    	return getField(emailBCCColIdx).replaceAll(";", ",");
	}

	public String getFileAttachValue() {
	    if (fileAttachColIdx == -1)
			return fileAttachValue;
	    else
			return excelEmailAttachFilesFolder + getField(fileAttachColIdx)  + excelEmailAttachFilesExt;
	}

	public String getField(int idx)
	{
		return (row.getCell(idx) == null ? "" : row.getCell(idx).getStringCellValue());
	}
	
	public void setSentFlag(String value)
	{
		cell = row.createCell(1);
		if (cell != null)
			cell.setCellValue(value);
	}
	
	public void writeChanges() throws IOException
	{
		String fileNameCopy = excelFilePath.replace(".xls", "-copy.xls");
	    os = new FileOutputStream(fileNameCopy);
		wb.write(os);
		os.close();
	}

}
