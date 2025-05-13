package may13;

import utilities.ExcelFileUtil;

public class Dummy {

	public static void main(String[] args) throws Throwable {
		ExcelFileUtil xl = new ExcelFileUtil("D:/myFile.xlsx");
		int rc = xl.rowCount("Emp");
		System.out.println(rc);
		for(int i=1;i<=rc;i++) {
			String fname =xl.getCellData("Emp", i, 0);
			String mname = xl.getCellData("Emp", i, 1);
			String lname = xl.getCellData("Emp", i, 2);
			String eid = xl.getCellData("Emp", i, 3);
			System.out.println(fname+"   "+mname+"  "+lname+"  "+eid);
			//xl.setCellData("Emp", i, 4, "Pass", "D:/sampleResults.xlsx");
			//xl.setCellData("Emp", i, 4, "Fail", "D:/sampleResults.xlsx");
			xl.setCellData("Emp", i, 4, "Blocked", "D:/sampleResults.xlsx");
		}

	}

}
