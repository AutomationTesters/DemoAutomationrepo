
public static void dataExtract() throws ClassNotFoundException, InvalidFormatException, IOException, SQLException {

		ExtentCucumberFormatter.stepTestThreadLocal.get().pass(" Fetching data from DB - Started");
		// prefixed index value variables with i.
		/*
		 * FileInputStream fis = new FileInputStream(new File( "DBInput/DBQuery.xlsx"));
		 */
		// Spread sheet loation and sheet name changed
		FileInputStream fis = new FileInputStream(new File("TestData/DBInputDetails.xlsx"));
		testDataQuery = new XSSFWorkbook(fis);
		testDataQueryConfigSheet = testDataQuery.getSheet("Config");
		testDataQueryQuerySheet = testDataQuery.getSheet("Query");
		testDataQueryConnDetailsSheet = testDataQuery.getSheet("ConnectionDetails");
		ExtentCucumberFormatter.stepTestThreadLocal.get()
				.pass(" Fetching data from DB output folder Location :" + "TestData/DBInputDetails.xlsx");
		int iExcelName = getColumnIndex(testDataQueryConfigSheet, "ExcelName"),
				iSheetName = getColumnIndex(testDataQueryConfigSheet, "SheetName"),
				iQuery = getColumnIndex(testDataQueryConfigSheet, "Query");

		// Code Begins

		for (int rowNum = 1; rowNum <= testDataQueryConfigSheet.getLastRowNum(); rowNum++) {
			if (testDataQueryConfigSheet.getRow(rowNum).getCell(iExcelName).getStringCellValue()
					.equalsIgnoreCase(ApplicationsVariables.scenarioName)) {
				/*
				 * System.out.println(testDataQueryConfigSheet.getRow(rowNum).getCell(
				 * iExcelName) .getStringCellValue());
				 */
				// Row(1) is used for Table name in Query Sheet
				ExtentCucumberFormatter.stepTestThreadLocal.get()
						.pass(" Fetching data from DB ScenarioName : " + ApplicationsVariables.scenarioName);
				//try {

					Connection connection = DriverManager.getConnection(
							"jdbc:oracle:thin:@//"
									+ testDataQueryConnDetailsSheet.getRow(2).getCell(0).getStringCellValue(),
							testDataQueryConnDetailsSheet.getRow(2).getCell(1).getStringCellValue(),
							testDataQueryConnDetailsSheet.getRow(2).getCell(2).getStringCellValue());
					/*
					 * System.out.println(testDataQueryConnDetailsSheet.getRow(2)
					 * .getCell(0).getStringCellValue() + " :" +
					 * testDataQueryConnDetailsSheet.getRow(2).getCell(1) .getStringCellValue() +
					 * " : " +testDataQueryConnDetailsSheet.getRow(2).getCell(2)
					 * .getStringCellValue() );
					 */
					// System.out.println("jdbc:oracle:thin:@//"+connectionString+","+userName+","+password);
					Statement statement = connection.createStatement();
					System.out.println("connection created");
					ResultSet resultSet = statement.executeQuery(
							testDataQueryConfigSheet.getRow(rowNum).getCell(iQuery).getStringCellValue().toString());
					System.out.println("resultset created");
					System.out.println(
							testDataQueryConfigSheet.getRow(rowNum).getCell(iQuery).getStringCellValue().toString());
					ExtentCucumberFormatter.stepTestThreadLocal.get().pass(" Fetching data from DB Query details: "
							+ testDataQueryConfigSheet.getRow(rowNum).getCell(iQuery).getStringCellValue().toString());
					for (int i = 1; i <= resultSet.getMetaData().getColumnCount(); i++) {
						resultSet.getMetaData().getColumnLabel(i);
					}
					excelName = testDataQueryConfigSheet.getRow(rowNum).getCell(iExcelName).getStringCellValue()
							.toString();
					String sheetName = testDataQueryConfigSheet.getRow(rowNum).getCell(iSheetName).getStringCellValue()
							.toString();
					 System.out.println("Excel Name :" + excelName);
					if (testDataQueryConfigSheet.getRow(rowNum).getCell(iExcelName).getStringCellValue()
							.equalsIgnoreCase("LBXNewCustomerExtractDB")) {
						writeDBResultToExcelLBXCust(resultSet);
					} else {
						writeDBResultToExcel(resultSet, excelName, sheetName);
					}
					statement.close();
					connection.close();
					ExtentCucumberFormatter.stepTestThreadLocal.get().pass(" Fetching data from DB - Completed");
				//} catch (NullPointerException e) {
				//	e.printStackTrace();
				//}
			}

		}

	}
private static void writeDBResultToExcel(ResultSet resultSet, String excelName, String sheetName)
			throws SQLException, FileNotFoundException, IOException, InvalidFormatException {

		// Result Workbook
		XSSFWorkbook resultWorkbook = null;
		File dbResultFolder = new File("DBResults/" + TestNGCukesRunner.testResultFolderNameDownload);
		if (!dbResultFolder.exists()) {
			dbResultFolder.mkdir();
		}
		File file = new File(dbResultFolder+"/" + excelName + ".xlsx");
		if (file.exists()) {
			  FileUtils.cleanDirectory(new File ("DBResults/" + TestNGCukesRunner.testResultFolderNameDownload));  
			 /* //File resultFile = new File(dbResultFolder+"/" + excelName + ".xlsx");
			resultWorkbook = new XSSFWorkbook(new FileInputStream(file));
			System.out.println("Inside IF");*/}
			  resultWorkbook = new XSSFWorkbook();		/*if(resultWorkbook.getSheetAt(0).getSheetName().equalsIgnoreCase(sheetName)) {
			
			System.out.println(sheetName+ " Already exists");}*/
		XSSFSheet resultSheet = resultWorkbook.createSheet(sheetName);
		int resultRowNum = 0;
		XSSFRow resultRow = resultSheet.createRow(resultRowNum);
		XSSFCell resultCell = null;

		ResultSetMetaData rsMetaData = resultSet.getMetaData();
		int columnCount = rsMetaData.getColumnCount();

		for (int i = 1; i <= columnCount; i++) {
			resultCell = resultRow.createCell(i - 1);
			// resultCell.setCellStyle(headingStyle(resultWorkbook));
			resultCell.setCellValue(resultSet.getMetaData().getColumnLabel(i));
			resultSheet.autoSizeColumn(resultCell.getColumnIndex());
		}

		while (resultSet.next()) {
			resultRow = resultSheet.createRow(++resultRowNum);
			for (int i = 1; i <= columnCount; i++) {
				String columnName = rsMetaData.getColumnLabel(i);
				Cell cell = resultRow.createCell(getColumnIndex(resultSheet, columnName));
				cell.setCellValue(resultSet.getString(columnName));
				resultSheet.autoSizeColumn(cell.getColumnIndex());
			}
		}

		resultWorkbook.write(new FileOutputStream(file));
		// resultWorkbook.close();
	}	
