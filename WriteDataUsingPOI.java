// import statements
public class WriteDataUsingPOI {

public void writeSchoolsData()
    {
//************************* CODE TO WRITE CELLS TO A SHEET ********************************
        // Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
  
        // Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("schools-sheet");
  
        // This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[]{ "ID", "SCHOOL NAME", "SCHOOL ADDRESS" });
        data.put("2", new Object[]{ 1, "Junior Model School", "Islamabad, Pakistan" });
        data.put("3", new Object[]{ 2, "Senior Secondary School", "Lahore, Pakistan" });
        data.put("4", new Object[]{ 3, "Post Grad Collage", "Islamabad, Pakistan" });
  
        // Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
            // this creates a new row in the sheet
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                // this line creates a cell in the next column of that row
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String)
                    cell.setCellValue((String)obj);
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
//******************* CODE TO WRITE THE FILE/WORKBOOK ON DISK *********************
        try {
            FileOutputStream out = new FileOutputStream(new File("preferred-schools.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("preferred-schools.xlsx written successfully on disk.");
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }
}