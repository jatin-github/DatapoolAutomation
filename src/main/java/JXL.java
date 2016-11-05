
import java.io.File;
import jxl.write.Label;
import java.io.IOException;
import java.security.Timestamp;

import com.sun.corba.se.impl.ior.GenericTaggedComponent;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.biff.formula.ParseContext;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class JXL {

	static Workbook SWBook;
	static Workbook DWBook;
	static Workbook MWBook;
	static Sheet SSheet;
	static Sheet DSheet;
	static Sheet MSheet;
	static WritableWorkbook TWBook = null;
	static WritableSheet TSheet = null;

	static int counter = 1;
	public static int noOfMapEntriesIncludingHeader;

	static final String MappingFileName = "mapping";
	static final String One2OneMappingSheetName = "1to1-Map";
	static String ruleBookName = "RPI-RuleBook";

	static File outputWorkbook;
	/*
	 * Also we are using JXL which means it works fine with 'xls' format only,
	 * so we can omit file extension name. These names are restricted, you can't
	 * change these names.
	 */

	public static void main(String[] args) throws Exception {

		boolean flag = true;

		/*
		 * isXlsExists(MappingFileName); isSheetExists(MappingFileName,
		 * One2OneMappingSheetName);
		 */

		try {
			MWBook = Workbook
					.getWorkbook(new File(System.getProperty("user.dir") + "\\files\\" + MappingFileName + ".xls"));
			MSheet = MWBook.getSheet(One2OneMappingSheetName);
		

		} catch (Exception e) {
			e.printStackTrace();
		}
		noOfMapEntriesIncludingHeader = MSheet.getRows();

		System.out.println("noOfMapEntriesIncludingHeader:==" + noOfMapEntriesIncludingHeader);
		System.out.println("--//>>>>>>>>>><<<<<<<<<<//--");

		String currentSourceWB = "";
		String currentSourceSh = "";
		String currentSourceCol = "";

		String currentDestinationWB = "";
		String currentDestinationSh = "";
		String currentDestinationCol = "";

		String nextDestinationWB = "";
		String nextDestinationSh = "";
		String nextDestinationCol = "";

		String listOfSourceValues[];
		String listOfDestinationValues[];

		while (counter < noOfMapEntriesIncludingHeader) {
			int currentdestinationRow = 1;
			System.out.println("currentSourceWB = " + MSheet.getCell(0, counter).getContents());
			System.out.println("currentSourceSh = " + MSheet.getCell(1, counter).getContents());
			System.out.println("currentSourceCol = " + MSheet.getCell(2, counter).getContents());
			System.out.println("listOfSourceValues = " + MSheet.getCell(3, counter).getContents());
			System.out.println("listOfDestinationValues = " + MSheet.getCell(4, counter).getContents());
			System.out.println("currentDestinationWB = " + MSheet.getCell(5, counter).getContents());
			System.out.println("currentDestinationSh = " + MSheet.getCell(6, counter).getContents());
			System.out.println("currentDestinationCol = " + MSheet.getCell(7, counter).getContents());
			System.out.println("noOfcopies = " + MSheet.getCell(8, counter).getContents());
			System.out.println("--//>>>>>>>>>><<<<<<<<<<//--");

			if (counter + 1 < noOfMapEntriesIncludingHeader) {
				nextDestinationWB=MSheet.getCell(5, counter + 1).getContents();
				System.out.println("nextDestinationWB = " + MSheet.getCell(5, counter + 1).getContents());
				System.out.println("nextDestinationSh = " + MSheet.getCell(6, counter + 1).getContents());
				System.out.println("nextDestinationCol = " + MSheet.getCell(7, counter + 1).getContents());
			}

			// below code verifies destination

			if (MSheet.getCell(5, counter).getContents().equalsIgnoreCase("")
					|| MSheet.getCell(6, counter).getContents().equalsIgnoreCase("")
					|| MSheet.getCell(7, counter).getContents().equalsIgnoreCase("")) {
				throw new Exception("destination is not specified in row # " + (counter + 1));
				// This row number is your 'xls' exact row number.
			}

			if (MSheet.getCell(0, counter).getContents().equalsIgnoreCase("")
					|| MSheet.getCell(1, counter).getContents().equalsIgnoreCase("")
					|| MSheet.getCell(2, counter).getContents().equalsIgnoreCase("")) {

				
				
				
				System.out.println("Source unavailable for in row # = " + counter);

				if (MSheet.getCell(3, counter).getContents().equals("")
						|| MSheet.getCell(3, counter).getContents().equals("$")) {
					throw new Exception("Incompatiblility between source and destination values at Row #:= " + counter);
				}
				
				
				String sourceValues[] = MSheet.getCell(3, counter).getContents().split(",");
				String destinationValues[] = MSheet.getCell(4, counter).getContents().split(",");
				int destinationColNo = DSheet.findCell(MSheet.getCell(7, counter).getContents()).getColumn();
				int noOfRowsInSource = SSheet.getRows();
				int noOfcopies = Integer.parseInt(MSheet.getCell(8, counter).getContents());

				for (int i = 1; i < noOfRowsInSource; i++) {
					String valueTobePopulatedInDestination = "";
					for (int j = 0; j < noOfcopies; j++) {
						Label Name1 = new Label(destinationColNo, currentdestinationRow + j, destinationValues[0]);
						TSheet.addCell(Name1);
					}
					currentdestinationRow = currentdestinationRow + noOfcopies;
				}

			} else {

				try {
					SWBook = Workbook.getWorkbook(new File(System.getProperty("user.dir") + "\\files\\"
							+ MSheet.getCell(0, counter).getContents() + ".xls"));
					SSheet = SWBook.getSheet(MSheet.getCell(1, counter).getContents());
					DWBook = Workbook.getWorkbook(new File(System.getProperty("user.dir") + "\\files\\"
							+ MSheet.getCell(5, counter).getContents() + ".xls"));
					DSheet = DWBook.getSheet(MSheet.getCell(6, counter).getContents());
				} catch (IOException e) {
					e.printStackTrace();
					break;
				}
				if (flag) {
					outputWorkbook = new File(System.getProperty("user.dir") + "\\files\\" + MSheet.getCell(5, counter).getContents() + "-New" + ".xls");
					TWBook = Workbook.createWorkbook(outputWorkbook, DWBook);
					TSheet = TWBook.getSheet(MSheet.getCell(6, counter).getContents());
					flag = false;
				}

				// (-,-)
				if (MSheet.getCell(3, counter).getContents().equals("-")
						&& MSheet.getCell(4, counter).getContents().equals("-")) {
					System.out.println("source and destination No mapping");
					int sourceColNo = SSheet.findCell(MSheet.getCell(2, counter).getContents()).getColumn();
					int destinationColNo = DSheet.findCell(MSheet.getCell(7, counter).getContents()).getColumn();
					int noOfRowsInSource = SSheet.getRows();
					int noOfcopies = Integer.parseInt(MSheet.getCell(8, counter).getContents());

					for (int i = 1; i < noOfRowsInSource; i++) {
						for (int j = 0; j < noOfcopies; j++) {
							System.out.println("currentdestinationRow:--" + currentdestinationRow);

							Label Name1 = new Label(destinationColNo, currentdestinationRow + j,
									SSheet.getCell(sourceColNo, i).getContents());
							TSheet.addCell(Name1);
						}
						currentdestinationRow = currentdestinationRow + noOfcopies;
					}
				}
				// (-,Values)
				else if (MSheet.getCell(3, counter).getContents().equals("-")
						&& !(MSheet.getCell(4, counter).getContents().equals("-"))) {
					throw new Exception("Incompatiblility between source and destination values at Row #:= " + counter);

				}
				// (values,-)
				else if (!(MSheet.getCell(3, counter).getContents().equals("-"))
						&& (MSheet.getCell(4, counter).getContents().equals("-"))) {
					throw new Exception("Incompatiblility between source and destination values at Row #:= " + counter);
				}
				// (Values,Values)
				else if (!(MSheet.getCell(3, counter).getContents().equals("-"))
						&& !(MSheet.getCell(4, counter).getContents().equals("-"))) {
					
					if((MSheet.getCell(3, counter).getContents().equals(""))
							&& (MSheet.getCell(4, counter).getContents().equals("")))
					{
						throw new Exception("Incompatiblility between source and destination values at Row #:= " + counter);
					}
					if((MSheet.getCell(3, counter).getContents().equals(""))
							&& (MSheet.getCell(4, counter).getContents().equals("$")))
					{
						throw new Exception("Incompatiblility between source and destination values at Row #:= " + counter);
					}
					if((MSheet.getCell(3, counter).getContents().equals("$"))
							&& (MSheet.getCell(4, counter).getContents().equals("")))
					{
						throw new Exception("Incompatiblility between source and destination values at Row #:= " + counter);
					}
					
					
					
					String sourceValues[] = MSheet.getCell(3, counter).getContents().split(",");
					String destinationValues[] = MSheet.getCell(4, counter).getContents().split(",");
					int sourceColNo = SSheet.findCell(MSheet.getCell(2, counter).getContents()).getColumn();
					int destinationColNo = DSheet.findCell(MSheet.getCell(7, counter).getContents()).getColumn();
					int noOfRowsInSource = SSheet.getRows();
					int noOfcopies = Integer.parseInt(MSheet.getCell(8, counter).getContents());

					for (int i = 1; i < noOfRowsInSource; i++) {
						String valueTobePopulatedInDestination = "";
						for (int j = 0; j < noOfcopies; j++) {
							System.out.println("currentdestinationRow:--" + currentdestinationRow);

							for (int k = 0; k < sourceValues.length; k++) {

								if (sourceValues[k].contentEquals("$")) {
									if (destinationValues[k].contentEquals("$")) {
										valueTobePopulatedInDestination = "";
									} else {
										valueTobePopulatedInDestination = destinationValues[k];
									}

									break;
								} else if (sourceValues[k]
										.contentEquals(SSheet.getCell(sourceColNo, i).getContents())) {
									if (destinationValues[k].contentEquals("$")) {
										valueTobePopulatedInDestination = "";
									} else {
										valueTobePopulatedInDestination = destinationValues[k];
									}
									break;
								}
							}

							Label Name1 = new Label(destinationColNo, currentdestinationRow + j,
									valueTobePopulatedInDestination);
							TSheet.addCell(Name1);
						}
						currentdestinationRow = currentdestinationRow + noOfcopies;
					}

				}
				
			}
			if (!(nextDestinationWB.equalsIgnoreCase(MSheet.getCell(5, counter).getContents()))) {
				System.out.println("----------New Destination File-----------");
				TWBook.write();
				TWBook.close();
				flag = true;
			}


			counter++;
			
		}
		TWBook.write();
		TWBook.close();

		/*
		 * getMapSet();
		 */

	}

	public static void getMapSet() throws BiffException, IOException, WriteException {

		boolean flag = true;
		String[] stemp = { "", "", "" };
		String[] dtemp = { "", "", "" };
		String dtempnext[] = { "", "", "" };

		String currentSourceWB = "";
		String currentSourceSh = "";
		String currentSourceCol = "";

		String currentDestinationWB = "";
		String currentDestinationSh = "";
		String currentDestinationCol = "";

		Sheet tempsheet = MWBook.getSheet(ruleBookName);

		while (counter < noOfMapEntriesIncludingHeader) {

			final File outputWorkbook;

			// ------------------------getCell(Column number,Row
			// number);---------------------//

			System.out.println("I am counter:--" + counter);
			stemp = MSheet.getCell(0, counter).getContents().split(":");
			dtemp = MSheet.getCell(1, counter).getContents().split(":");
			System.out.println("MSheet.getRows()>>>." + MSheet.getRows());
			System.out.println(stemp[0]);
			System.out.println(stemp[1]);
			System.out.println(stemp[2] + "\n");
			System.out.println(dtemp[0]);
			System.out.println(dtemp[1]);
			System.out.println(dtemp[2] + "\n");

			SWBook = Workbook.getWorkbook(new File(System.getProperty("user.dir") + "\\files\\" + stemp[0] + ".xls"));
			SSheet = SWBook.getSheet(stemp[1]);
			DWBook = Workbook.getWorkbook(new File(System.getProperty("user.dir") + "\\files\\" + dtemp[0] + ".xls"));
			DSheet = DWBook.getSheet(dtemp[1]);

			if (counter + 1 < MSheet.getRows()) {

				dtempnext = MSheet.getCell(1, counter + 1).getContents().split(":");
				System.out.println("dtempnext[0]>>>>" + dtempnext[0]);
				System.out.println("dtempnext[1]>>>>" + dtempnext[1]);
				System.out.println("dtempnext[2]>>>>" + dtempnext[2] + "\n");

			}

			if (flag) {
				outputWorkbook = new File(System.getProperty("user.dir") + "\\files\\" + dtemp[0] + "-New" + ".xls");
				TWBook = Workbook.createWorkbook(outputWorkbook, DWBook);
				flag = false;
			}

			Cell[] destinationCol = tempsheet.getColumn(3);

			int currentmappedRcol = tempsheet.findCell(dtemp[2]).getColumn();
			int currentMappedRow = tempsheet.findCell(dtemp[2]).getRow();

			System.out.println("currentmappedRcol>>>" + currentmappedRcol);
			System.out.println("currentMappedRow>>>" + currentMappedRow);
			String[] dMap = tempsheet.getCell(currentmappedRcol - 1, currentMappedRow).getContents().split(",");
			String[] sMap = tempsheet.getCell(currentmappedRcol - 2, currentMappedRow).getContents().split(",");
			System.out.println("length " + sMap.length);
			System.out.println("value is " + sMap[0]);
			System.out.println("eqiivalent" + sMap[0].equalsIgnoreCase(""));

			/*
			 * System.out.pr intln(sMap[0]); System.out.println(sMap[1]);
			 * 
			 * System.out.println(dMap[0]); System.out.println(dMap[1]);
			 */
			TSheet = TWBook.getSheet(dtemp[1]);
			int sColNo = SSheet.findCell(stemp[2]).getColumn();
			int dColNo = DSheet.findCell(dtemp[2]).getColumn();
			int sRowNo = SSheet.getRows();

			int currentRowNo = 1;

			while (currentRowNo < sRowNo) {

				System.out.println("S value:- " + SSheet.getCell(sColNo, currentRowNo).getContents());
				int mapCounter = 0;

				while (!(mapCounter == sMap.length)) {

					if (SSheet.getCell(sColNo, currentRowNo).getContents().contentEquals(sMap[mapCounter])) {
						break;
					} else {
						mapCounter++;
					}
				}
				if (sMap[0].equalsIgnoreCase("-")) {
					Label Name1 = new Label(dColNo, currentRowNo, SSheet.getCell(sColNo, currentRowNo).getContents());
					TSheet.addCell(Name1);
				} else if (dMap[mapCounter].contentEquals("$")) {
					Label Name1 = new Label(dColNo, currentRowNo, " ");
					TSheet.addCell(Name1);
				} else {
					Label Name1 = new Label(dColNo, currentRowNo, dMap[mapCounter]);
					TSheet.addCell(Name1);

				}

				currentRowNo++;
			}
			System.out.println(currentRowNo);
			counter++;

			if (!(dtempnext[0].equalsIgnoreCase(dtemp[0]))) {
				System.out.println("----------New Destination File-----------");
				TWBook.write();
				TWBook.close();
				flag = true;
			}

		}
		TWBook.write();
		TWBook.close();
	}

	public static void isXlsExists(String xlsFileName) throws BiffException, IOException {
		try {
			Workbook temp = Workbook
					.getWorkbook(new File(System.getProperty("user.dir") + "\\files\\" + xlsFileName + ".xls"));
			temp.close();
		} catch (java.io.FileNotFoundException e) {
			e = new java.io.FileNotFoundException(
					"\n\nNo file found with the Name:- " + "'" + xlsFileName + "'" + "\n\n");
			e.printStackTrace();
		}
	}

	public static void isSheetExists(String xlsFileName, String sheetName) throws Exception {
		Workbook tempXls;
		Sheet tempsheet;
		Exception e;
		isXlsExists(xlsFileName);
		tempXls = Workbook.getWorkbook(new File(System.getProperty("user.dir") + "\\files\\" + xlsFileName + ".xls"));
		tempsheet = tempXls.getSheet(sheetName);
		if (tempXls.getSheet(sheetName) == null) {

			throw e = new Exception("\n\nNo sheet found with the Name:- " + "'" + sheetName + "'" + "\n\n");
		}
		tempXls.close();

	}

}