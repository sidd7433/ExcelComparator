package org.excelcomparator;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Assert;
import org.junit.Test;

public class ExcelComparatorTest {

	private static final String RESOURCE_A1_XLS = "src/test/resources/testfiles/A1.xls";
	private static final String RESOURCE_A1_XLSX = "src/test/resources/testfiles/A1.xlsx";
	private static final String RESOURCE_A2_XLSX = "src/test/resources/testfiles/A2.xlsx";
	private static final String RESOURCE_B1_XLS = "src/test/resources/testfiles/B1.xls";
	private static final String RESOURCE_B1_XLSX = "src/test/resources/testfiles/B1.xlsx";
	private static final String RESOURCE_B2_XLSX = "src/test/resources/testfiles/B2.xlsx";
	private static final String RESOURCES_A2_XLS = "src/test/resources/testfiles/A2.xls";
	private static final String RESOURCES_B2_XLS = "src/test/resources/testfiles/B2.xls";

	@Test
	public void testDifferentFileOldFormat() throws FileNotFoundException {
		final File file1 = new File(ExcelComparatorTest.RESOURCES_A2_XLS);
		final File file2 = new File(ExcelComparatorTest.RESOURCES_B2_XLS);

		Assert.assertFalse("Test to Compare Different Files", ExcelComparator.compareWithoutDetails(file1, file2));
	}

	@Test
	public void testDifferentFiles() throws FileNotFoundException {
		final File file1 = new File(ExcelComparatorTest.RESOURCE_A2_XLSX);
		final File file2 = new File(ExcelComparatorTest.RESOURCE_B2_XLSX);

		Assert.assertFalse("Test to Compare Different Files", ExcelComparator.compareWithoutDetails(file1, file2));
	}

	@Test
	public void testDifferentFilesWithDetails() throws EncryptedDocumentException, InvalidFormatException, IOException {
		final File file1 = new File(ExcelComparatorTest.RESOURCE_A2_XLSX);
		final File file2 = new File(ExcelComparatorTest.RESOURCE_B2_XLSX);
		final ExcelFileDifference expectedObj = ExcelComparator.compare(file1, file2);
		Assert.assertTrue("Differences Found", expectedObj.isDifferenceFound);
		Assert.assertEquals("Compare Size", "1", expectedObj.listOfDifferences.size() + "");
		Assert.assertTrue("Cell content ", expectedObj.listOfDifferences.get(0)
				.contains("Sheet1 -> A1 [abcd] != workbook2 -> Sheet1 -> A1 [abcde]"));
	}

	@Test
	public void testFile1NotFound() {
		final File file1 = new File("src/test/resources/testfiles/A.xlsx");
		final File file2 = new File(ExcelComparatorTest.RESOURCE_B2_XLSX);

		Assert.assertFalse("Test to File 1 Exist", ExcelComparator.compareWithoutDetails(file1, file2));
	}

	@Test
	public void testFile1NotFoundForDetails() {
		final File file1 = new File("src/test/resources/testfiles/A.xlsx");
		final File file2 = new File(ExcelComparatorTest.RESOURCE_B2_XLSX);
		final ExcelFileDifference excelFileDifference = ExcelComparator.compare(file1, file2);
		Assert.assertNotNull("Test to File 1 Exist", excelFileDifference.error);
	}

	@Test
	public void testFile2NotFound() {
		final File file1 = new File(ExcelComparatorTest.RESOURCE_A1_XLSX);
		final File file2 = new File("src/test/resources/testfiles/B.xlsx");

		Assert.assertFalse("Test to File 1 Exist", ExcelComparator.compareWithoutDetails(file1, file2));
	}

	@Test
	public void testFile2NotFoundForDetails() {
		final File file1 = new File(ExcelComparatorTest.RESOURCE_A1_XLSX);
		final File file2 = new File("src/test/resources/testfiles/B.xlsx");

		final ExcelFileDifference excelFileDifference = ExcelComparator.compare(file1, file2);
		Assert.assertNotNull("Test to File 2 Exist", excelFileDifference.error);
	}

	@Test
	public void testSameFiles() throws FileNotFoundException {
		final File file1 = new File(ExcelComparatorTest.RESOURCE_A1_XLSX);
		final File file2 = new File(ExcelComparatorTest.RESOURCE_B1_XLSX);

		Assert.assertTrue("Test to Compare same Files", ExcelComparator.compareWithoutDetails(file1, file2));
	}

	@Test
	public void testSameFilesOldFormat() throws FileNotFoundException {
		final File file1 = new File(ExcelComparatorTest.RESOURCE_A1_XLS);
		final File file2 = new File(ExcelComparatorTest.RESOURCE_B1_XLS);

		Assert.assertTrue("Test to Compare same Files", ExcelComparator.compareWithoutDetails(file1, file2));
	}

	@Test
	public void testSameFilesWithDetails() throws EncryptedDocumentException, InvalidFormatException, IOException {
		final File file1 = new File(ExcelComparatorTest.RESOURCE_A1_XLSX);
		final File file2 = new File(ExcelComparatorTest.RESOURCE_B1_XLSX);
		final ExcelFileDifference expectedObj = ExcelComparator.compare(file1, file2);
		Assert.assertFalse(expectedObj.isDifferenceFound);
		Assert.assertEquals("0", expectedObj.listOfDifferences.size() + "");
	}

}
