
package Service;

import DAOduan.DAOkehoach;
import DAOduan.DAOkehoachthiimprements;
import Model.Sinhvien;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class checksv {

    private DAOkehoach ds;
    private ArrayList<Sinhvien> lstsv;
    private ArrayList<Sinhvien> lstthi;
    private ArrayList<Sinhvien> lstcamthi;

    public checksv() {
        this.ds = new DAOkehoachthiimprements();
        this.lstsv = new ArrayList<>();
        this.lstcamthi = new ArrayList<>();
        this.lstthi = new ArrayList<>();
    }
    
    private XSSFSheet createSheet(String namefile) throws Exception {
        FileInputStream excel = new FileInputStream(namefile);
        XSSFWorkbook workbook = new XSSFWorkbook(excel);
        XSSFSheet sheet = workbook.getSheetAt(0);
        return sheet;
    }

    private Iterator createiterator(XSSFSheet sheet) throws Exception {
        Iterator<Row> iterator = sheet.iterator();
        return iterator;
    }

    public void ktramondauvao(String namefile) throws Exception {
        List<Integer> dscolumndiem = new ArrayList<>();
        XSSFSheet sheet = createSheet(namefile);
        Iterator<Row> iterator = createiterator(sheet);
        sheet.getRow(6).forEach(cell -> {
            if (cell.getStringCellValue().equalsIgnoreCase("Bài học online")) {
                sheet.getRow(6).forEach(cellonl -> {
                    if (cellonl.getStringCellValue().equalsIgnoreCase("MSSV") || cellonl.getStringCellValue().equalsIgnoreCase("Họ và tên")
                            || cellonl.getStringCellValue().equalsIgnoreCase("Bài học online") || cellonl.getStringCellValue().equalsIgnoreCase("Trạng thái")) {
                        dscolumndiem.add(cellonl.getColumnIndex());
                    }
                });
                try {
                    lstsv = ds.docexcelloai1(iterator, dscolumndiem);
                    checkdiemonl(lstsv);
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
        });

        if (lstsv.size() == 0) {
            sheet.getRow(6).forEach(cellquiz -> {
                if (cellquiz.getStringCellValue().equalsIgnoreCase("MSSV") || cellquiz.getStringCellValue().equalsIgnoreCase("Họ và tên")
                        || cellquiz.getStringCellValue().equalsIgnoreCase("Quiz 1") || cellquiz.getStringCellValue().equalsIgnoreCase("Quiz 2")
                        || cellquiz.getStringCellValue().equalsIgnoreCase("Quiz 3") || cellquiz.getStringCellValue().equalsIgnoreCase("Quiz 4")
                        || cellquiz.getStringCellValue().equalsIgnoreCase("Quiz 5") || cellquiz.getStringCellValue().equalsIgnoreCase("Quiz 6")
                        || cellquiz.getStringCellValue().equalsIgnoreCase("Quiz 7") || cellquiz.getStringCellValue().equalsIgnoreCase("Quiz 8")
                        || cellquiz.getStringCellValue().equalsIgnoreCase("Trạng thái")) {
                    dscolumndiem.add(cellquiz.getColumnIndex());
                }
            });
            try {
                lstsv = ds.docexcelloai2(iterator, dscolumndiem);
                checkdiemquiz(lstsv);
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
    }

    public void docfilediemdanh(String namefile) throws Exception {
        List<Integer> dscolumndd = new ArrayList<>();
        XSSFSheet sheet = createSheet(namefile);
        Iterator<Row> iterator = createiterator(sheet);
                sheet.getRow(6).forEach(cellonl -> {
                    if (cellonl.getStringCellValue().equalsIgnoreCase("MSSV") || cellonl.getStringCellValue().equalsIgnoreCase("Họ và tên") 
                            || cellonl.getStringCellValue().equalsIgnoreCase("Trạng thái")) {
                        dscolumndd.add(cellonl.getColumnIndex());
                    }
                });
        lstsv=ds.docexceldiemdanh(iterator, dscolumndd);
        checkmondiemdanh(lstsv);
    }

    private void checkdiemonl(ArrayList<Sinhvien> lst) {
        for (int i = 0; i < this.lstsv.size(); i++) {
            if (lstsv.get(i).getDiemonl() < 7.5 || lstsv.get(i).getTinhtrang().equalsIgnoreCase("Attendance failed")) {
                lstcamthi.add(new Sinhvien(lstsv.get(i).getDiemonl(), lstsv.get(i).getTensv(), lstsv.get(i).getMasv(), "cấm thi"));
            } else {
                lstthi.add(new Sinhvien(lstsv.get(i).getDiemonl(), lstsv.get(i).getTensv(), lstsv.get(i).getMasv(), ""));
            }
        }
    }

    private void checkdiemquiz(ArrayList<Sinhvien> ds) {
        for (int i = 0; i < this.lstsv.size(); i++) {
            if (lstsv.get(i).getDiemonl() < 80 || lstsv.get(i).getTinhtrang().equalsIgnoreCase("Attendance failed")) {
                lstcamthi.add(new Sinhvien(lstsv.get(i).getDiemonl(), lstsv.get(i).getTensv(), lstsv.get(i).getMasv(), "cấm thi"));
            } else {
                lstthi.add(new Sinhvien(lstsv.get(i).getDiemonl(), lstsv.get(i).getTensv(), lstsv.get(i).getMasv(), ""));
            }
        }
    }

    private void checkmondiemdanh(ArrayList<Sinhvien> ds) {
        for (int i = 0; i < this.lstsv.size(); i++) {
            if (lstsv.get(i).getTinhtrang().equalsIgnoreCase("Attendance failed")) {
                lstcamthi.add(new Sinhvien(lstsv.get(i).getDiemonl(), lstsv.get(i).getTensv(), lstsv.get(i).getMasv(), "cấm thi"));
            } else {
                lstthi.add(new Sinhvien(lstsv.get(i).getDiemonl(), lstsv.get(i).getTensv(), lstsv.get(i).getMasv(), ""));
            }
        }
        System.out.println(lstthi.size()+"abc");
    }

    public void xuatdssthi(String namefile, String block, int count) throws Exception {
        ds.xuatdssthifileword(namefile + ".docx", count, lstthi);
        ds.xuatlichthi(namefile + ".xlsx", block, count, lstthi, lstcamthi);
    }

    public int xuatdssvthi() {
        return this.lstthi.size();
    }

}
