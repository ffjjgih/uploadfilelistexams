
package DAOduan;

import Model.Inputkehoachthi;
import Model.Sinhvien;
import Sqlserver.Connect;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public class DAOkehoachthiimprements implements DAOkehoach {

    private ArrayList<Inputkehoachthi> lstkht;
    private ArrayList<Sinhvien> lstsv;
    String lop = null, mamon, ma1, ma, mua;
    private static Connect getconnect;
    private Cell cell;
    private ArrayList<Integer> lstrow = new ArrayList<>();
    private ArrayList<Integer> lstcountsv = new ArrayList<>();

    public static Connection getconnection() throws Exception {
        Connection con = getconnect.getConnection();
        return con;
    }

    public DAOkehoachthiimprements() {
        this.lstkht = new ArrayList<>();
        this.lstsv = new ArrayList<>();
    }

    @Override
    public void docexcel(String namefile) throws Exception {
        try {
            String mamon, mamonhoc, phongthi, ma;
            java.util.Date ngay = null;
            int cathi;
            FileInputStream excel = new FileInputStream(namefile);
            XSSFWorkbook workbook = new XSSFWorkbook(excel);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = sheet.iterator();
            while (iterator.hasNext()) {
                Row row = iterator.next();
                if (row.getRowNum() > 131) {
                    mamonhoc = row.getCell(6).getStringCellValue();
                    phongthi = row.getCell(3).getStringCellValue();
                    ngay = row.getCell(1).getDateCellValue();
                    cathi = (int) row.getCell(2).getNumericCellValue();
                    mamon = row.getCell(10).getStringCellValue();
                    if (mamonhoc.length() > 0 && phongthi.length() > 0 && cathi > 0 && mamon.length() > 0) {
                        ma = mamon.replace(mamon.substring(mamon.length() - 6), "");
                        lstkht.add(new Inputkehoachthi(0, cathi, ma, phongthi, mamonhoc, ngay));
                    }

                }
            }
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        }

    }

    @Override
    public void luudb() throws Exception {
        try {
            String sql = "insert into ke_hoach_thi values(?,?,?,?,?)";
            PreparedStatement ps = getconnection().prepareStatement(sql);
            for (Inputkehoachthi x : this.lstkht) {
                java.util.Date day = x.getNgaythi();
                java.sql.Date ngay = new java.sql.Date(day.getTime());
                ps.setDate(1, ngay);
                ps.setString(2, x.getPhongthi());
                ps.setInt(3, x.getCathi());
                ps.setString(4, x.getMamon());
                ps.setString(5, x.getLop());
                ps.addBatch();
            }
            getconnection().commit();
            ps.executeBatch();
            getconnection().close();
            ps.close();
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
    }

    @Override
    public ArrayList<Sinhvien> docexcelloai1(Iterator<Row> iterquiz,List<Integer> lstcell) throws Exception {
        try {
            while (iterquiz.hasNext()) {
                Row row = iterquiz.next();
                Sinhvien mh1 = new Sinhvien();
                Iterator<Cell> celliter = row.cellIterator();
                if (row.getRowNum() == 2) {
                    //lop = row.getCell(3).getStringCellValue();
                    lop = "com108.1";
                    System.out.println(lop);
                }
                if (row.getRowNum() == 3) {
                    mamon = row.getCell(3).getStringCellValue();
                    System.out.println(mamon);
                }
                if (row.getRowNum() == 4) {
                    mua = row.getCell(3).getStringCellValue();
                }
                if (row.getRowNum() > 7) {
                    while (celliter.hasNext()) {

                        cell = celliter.next();
                        if (cell.getColumnIndex() == lstcell.get(0)) {
                            mh1.setMasv(row.getCell(lstcell.get(0)).getStringCellValue());
                        }
                        if (cell.getColumnIndex() == lstcell.get(1)) {
                            mh1.setTensv(row.getCell(lstcell.get(1)).getStringCellValue());
                        }
                        if (cell.getColumnIndex() == lstcell.get(2)) {
                            mh1.setDiemonl((double) row.getCell(lstcell.get(2)).getNumericCellValue());
                        }
                        if (cell.getColumnIndex() == lstcell.get(3)) {
                            mh1.setTinhtrang(row.getCell(lstcell.get(3)).getStringCellValue());
                        }
                    }
                    lstsv.add(mh1);
                }
            }
            xuatkehoachthi(lop, mamonhoc(mamon));
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        }
        return this.lstsv;
    }


    @Override
    public void xuatlichthi(String namefile, String lanthi, int count,ArrayList<Sinhvien> svthi,ArrayList<Sinhvien> svcthi) throws Exception {
        FileOutputStream fos = new FileOutputStream(namefile);
        XSSFWorkbook xssfw = new XSSFWorkbook();
        XSSFRow row, row1, row2, row3, row4, row5, row6;
        XSSFCell cellB, cellC, cellD, cellE, cellF, cellG, cellH;
        ArrayList<Sinhvien> lstsvkdt = new ArrayList<>();
        lstsvkdt.addAll(svcthi);
        XSSFSheet sheet = xssfw.createSheet(ngaythi(lstkht.get(0).getNgaythi()));
        int vitisv, slsv1cathi;
        for (int j = 0; j <= count; j++) {
            xuatbuoithi(j, count,svthi);
            if (j < count && j > 0) {
                sheet = xssfw.createSheet(ngaythi(lstkht.get(j).getNgaythi()));
            }
            if (count == j) {
                sheet = xssfw.createSheet("cấm thi");
            }
            XSSFCellStyle style = xssfw.createCellStyle();
            style.setBorderTop(BorderStyle.MEDIUM);
            style.setBorderBottom(BorderStyle.MEDIUM);
            style.setBorderLeft(BorderStyle.MEDIUM);
            style.setBorderRight(BorderStyle.MEDIUM);
            sheet.setColumnWidth(2, 6000);
            sheet.setColumnWidth(3, 3000);
            sheet.setColumnWidth(0, 1000);
            row1 = sheet.createRow((short) 0);
            row2 = sheet.createRow((short) 1);
            row3 = sheet.createRow((short) 2);
            row4 = sheet.createRow((short) 3);
            row5 = sheet.createRow((short) 4);
            row6 = sheet.createRow((short) 6);
            cellB = row6.createCell((short) 0);
            cellB.setCellStyle(style);
            cellB.setCellValue("TT");
            cellC = row6.createCell((short) 1);
            cellC.setCellStyle(style);
            cellC.setCellValue("MSSSV");
            cellD = row6.createCell((short) 2);
            cellD.setCellStyle(style);
            cellD.setCellValue("Họ Tên");
            cellE = row6.createCell((short) 3);
            cellE.setCellStyle(style);
            cellE.setCellValue("Lớp");
            cellF = row6.createCell((short) 4);
            cellF.setCellStyle(style);
            cellF.setCellValue("Kí Tên");
            cellG = row6.createCell((short) 5);
            cellG.setCellStyle(style);
            cellG.setCellValue("Điểm");
            cellH = row6.createCell((short) 6);
            cellH.setCellStyle(style);
            cellH.setCellValue("Ghi chú");
            cellB = row1.createCell((short) 3);
            cellB.setCellValue("Danh Sách Sinh Viên Thi");
            cellB = row2.createCell((short) 3);
            cellB.setCellValue("Kỳ:" + mua);
            cellB = row3.createCell((short) 3);
            cellB.setCellValue("Môn Thi:" + mamon);
            cellB = row4.createCell((short) 3);
            if (j < count) {
                cellB.setCellValue("phòng Thi:" + lstkht.get(j).getPhongthi());
                cellB = row5.createCell((short) 0);
                cellB.setCellValue("NGày Thi: " + lstkht.get(j).getNgaythi());
                cellC = row5.createCell((short) 3);
                cellC.setCellValue("giờ thi: " + lstkht.get(j).giothi());
                cellD = row5.createCell((short) 6);
                cellD.setCellValue("Lần Thi: " + lanthi);
            }
            if (j < count) {
                slsv1cathi = lstrow.get(j);
                vitisv = lstcountsv.get(j);
                for (int i = 1; i < slsv1cathi; i++) {
                    row = sheet.createRow((short) i + 6);
                    cellB = row.createCell((short) 0);
                    cellB.setCellStyle(style);
                    cellB.setCellValue(i);
                    cellC = row.createCell((short) 1);
                    cellC.setCellStyle(style);
                    cellC.setCellValue(svthi.get(i + vitisv - 1).getMasv());
                    cellD = row.createCell((short) 2);
                    cellD.setCellStyle(style);
                    cellD.setCellValue(svthi.get(i + vitisv - 1).getTensv());
                    cellE = row.createCell((short) 3);
                    cellE.setCellStyle(style);
                    cellE.setCellValue(lop.replaceAll("\\s+", ""));
                    cellF = row.createCell((short) 4);
                    cellF.setCellStyle(style);
                    cellG = row.createCell((short) 5);
                    cellG.setCellStyle(style);
                    cellH = row.createCell((short) 6);
                    cellH.setCellStyle(style);
                }
            }

            if (count == j) {
                for (int i = 0; i < svcthi.size(); i++) {
                    row = sheet.createRow((short) i + 7);
                    cellB = row.createCell((short) 0);
                    cellB.setCellValue(i + 1);
                    cellC = row.createCell((short) 1);
                    cellC.setCellValue(svcthi.get(i).getMasv());
                    cellD = row.createCell((short) 2);
                    cellD.setCellValue(svcthi.get(i).getTensv());
                    cellE = row.createCell((short) 3);
                    cellE.setCellValue(lop);
                    cellF = row.createCell((short) 6);
                    cellF.setCellValue(svcthi.get(i).getTinhtrang());
                }
            }
        }
        xssfw.write(fos);
        xssfw.close();
        fos.close();
    }

    @Override
    public void xuatkehoachthi(String lop, String ma) throws Exception {
        try {
            Statement ps = getconnect.getConnection().createStatement();
            String SQL = "SELECT * FROM ke_hoach_thi where mamon='" + ma + "' and lop='" + lop + "' order by ngaythi";
            ResultSet rs = ps.executeQuery(SQL);
            while (rs.next()) {
                Inputkehoachthi v = new Inputkehoachthi();
                v.setId(rs.getInt("ID"));
                v.setNgaythi(rs.getDate("NGAYTHI"));
                v.setPhongthi(rs.getString("PHONGTHI"));
                v.setCathi(rs.getInt("CATHI"));
                v.setMamon(rs.getString("MAMON"));
                v.setLop(rs.getString("LOP").replaceAll("\\s+", ""));
                this.lstkht.add(v);
            }
            ps.close();
            getconnect.getConnection().close();
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
    }

    @Override
    public ArrayList<Sinhvien> docexcelloai2(Iterator<Row> iterquiz,List<Integer> lstcell) throws Exception {
        try {
            double countquiz=0;
            while (iterquiz.hasNext()) {
                Row row = iterquiz.next();
                Sinhvien mh1 = new Sinhvien();
                Iterator<Cell> celliter = row.cellIterator();
                if (row.getRowNum() == 2) {
                    //mamon = row.getCell(3).getStringCellValue();
                    lop = "com108.1";
                }
                if (row.getRowNum() == 3) {
                    //lop = row.getCell(3).getStringCellValue();
                    mamon=" (com108)";
                }
                if (row.getRowNum() == 4) {
                    mua = row.getCell(3).getStringCellValue();
                }
                if (row.getRowNum() > 7) {
                    while(celliter.hasNext()){
                        cell = celliter.next();
                        if (cell.getColumnIndex() == lstcell.get(10)) {
                            mh1.setTinhtrang(row.getCell(lstcell.get(10)).getStringCellValue());
                        }
                        if (cell.getColumnIndex() == lstcell.get(0)) {
                            mh1.setMasv(row.getCell(lstcell.get(0)).getStringCellValue());
                        }
                        if (cell.getColumnIndex() == lstcell.get(1)) {
                            mh1.setTensv(row.getCell(lstcell.get(1)).getStringCellValue());
                        }
                        if (cell.getColumnIndex() == lstcell.get(2)) {
                            countquiz+=(double) row.getCell(lstcell.get(2)).getNumericCellValue();
                        }
                        if (cell.getColumnIndex() == lstcell.get(3)) {
                            countquiz+=(double) row.getCell(lstcell.get(3)).getNumericCellValue();
                        }
                        if (cell.getColumnIndex() == lstcell.get(4)) {
                            countquiz+=(double) row.getCell(lstcell.get(4)).getNumericCellValue();
                        }
                        if (cell.getColumnIndex() == lstcell.get(5)) {
                            countquiz+=(double) row.getCell(lstcell.get(5)).getNumericCellValue();
                        }
                        if (cell.getColumnIndex() == lstcell.get(6)) {
                            countquiz+=(double) row.getCell(lstcell.get(6)).getNumericCellValue();
                        }
                        if (cell.getColumnIndex() == lstcell.get(7)) {
                            countquiz+=(double) row.getCell(lstcell.get(7)).getNumericCellValue();
                        }
                        if (cell.getColumnIndex() == lstcell.get(8)) {
                            countquiz+=(double) row.getCell(lstcell.get(8)).getNumericCellValue();
                        }
                        if (cell.getColumnIndex() == lstcell.get(9)) {
                            countquiz+=(double) row.getCell(lstcell.get(9)).getNumericCellValue();
                        }
                    }
                    mh1.setDiemonl(countquiz);
                    lstsv.add(mh1);
                    countquiz=0;
                }
            }
            xuatkehoachthi(lop, mamonhoc(mamon));
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        }
        return this.lstsv;
    }


    private String mamonhoc(String id) {
        //String idmon = id.substring(id.lastIndexOf(' ') + 1);
        //String idmh = id.substring(1, id.length() - 1);
        String idmh = id.substring(2, id.length() - 1);
        return idmh;
    }

    private String ngaythi(java.util.Date ngay) {
        SimpleDateFormat sdf = new SimpleDateFormat("EEE, MMM d, yyyy");
        String ngayDate = sdf.format(ngay);
        return ngayDate;
    }

    @Override
    public void xuatdssthifileword(String namefile, int count,ArrayList<Sinhvien> svthi) throws Exception {
        FileOutputStream fos = new FileOutputStream(namefile);
        XWPFDocument xwpfd = new XWPFDocument();

        XWPFTable table;
        int vitisv, slsv1cathi;
        for (int j = 0; j < count; j++) {
            xuatbuoithi(j, count,svthi);
            XWPFParagraph tille = xwpfd.createParagraph();
            tille.setAlignment(ParagraphAlignment.CENTER);
            String title = "PHIẾU CHẤM THỰC HÀNH/BẢO VỆ ASSIGNMENT CUỐI MÔN HỌC";
            XWPFRun titleRun = tille.createRun();
            titleRun.setBold(true);
            titleRun.setFontFamily("Times New Roman");
            titleRun.setText(title);
            titleRun.setFontSize(14);
            slsv1cathi = lstrow.get(j);
            vitisv = lstcountsv.get(j);
            table = xwpfd.createTable(slsv1cathi, 9);
            table.getRow(0).setHeight(700);
            table.getRow(0).getCell(0).setWidth("500");
            table.getRow(0).getCell(1).setWidth("1500");
            table.getRow(0).getCell(2).setWidth("2500");
            table.getRow(0).getCell(3).setWidth("1700");
            table.getRow(0).getCell(4).setWidth("1700");
            table.getRow(0).getCell(5).setWidth("1700");
            table.getRow(0).getCell(6).setWidth("1000");
            table.getRow(0).getCell(7).setWidth("1000");
            table.getRow(0).getCell(8).setWidth("1500");
            table.getRow(0).getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.BOTTOM);
            table.getRow(0).getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.BOTTOM);
            table.getRow(0).getCell(2).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.BOTTOM);
            table.getRow(0).getCell(3).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.BOTTOM);
            table.getRow(0).getCell(4).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.BOTTOM);
            table.getRow(0).getCell(5).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.BOTTOM);
            table.getRow(0).getCell(6).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.BOTTOM);
            table.getRow(0).getCell(7).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.BOTTOM);
            table.getRow(0).getCell(8).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.BOTTOM);
            table.getRow(0).getCell(0).setText("TT");
            table.getRow(0).getCell(1).setText("MASV");
            table.getRow(0).getCell(2).setText("Họ Tên");
            table.getRow(0).getCell(3).setText("G3");
            table.getRow(0).getCell(4).setText("G4-G5");
            table.getRow(0).getCell(5).setText("G6");
            table.getRow(0).getCell(6).setText("Điểm bảo vệ");
            table.getRow(0).getCell(7).setText("SV ký nhận");
            table.getRow(0).getCell(8).setText("Nhận xét");
            for (int i = 1; i < slsv1cathi; i++) {
                table.getRow(i).setHeight(1300);
                table.getRow(i).getCell(0).setWidth("500");
                table.getRow(i).getCell(1).setWidth("1500");
                table.getRow(i).getCell(2).setWidth("2500");
                table.getRow(i).getCell(3).setWidth("1700");
                table.getRow(i).getCell(4).setWidth("1700");
                table.getRow(i).getCell(5).setWidth("1700");
                table.getRow(i).getCell(6).setWidth("1000");
                table.getRow(i).getCell(7).setWidth("1000");
                table.getRow(i).getCell(8).setWidth("1500");
                table.getRow(i).getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                table.getRow(i).getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                table.getRow(i).getCell(2).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                table.getRow(i).getCell(0).setText(i + "");
                table.getRow(i).getCell(1).setText(svthi.get(vitisv + i - 1).getMasv());
                table.getRow(i).getCell(2).setText(svthi.get(vitisv + i - 1).getTensv());
                table.getRow(i).getCell(3).setText("");
                table.getRow(i).getCell(4).setText("");
                table.getRow(i).getCell(5).setText("");
                table.getRow(i).getCell(6).setText("");
                table.getRow(i).getCell(7).setText("");
                table.getRow(i).getCell(8).setText("");

            }
            tille.setPageBreak(true);
        }

        xwpfd.write(fos);
        xwpfd.close();
        fos.close();
    }

    private void xuatbuoithi(int i, int count,ArrayList<Sinhvien> svddk) {
        count=count-1;
        if (i == 0) {
            if (svddk.size() == 25 || svddk.size() == 26 || svddk.size() > 36) {
                lstcountsv.add(0);
                lstrow.add(14);
            } else {
                lstcountsv.add(0);
                lstrow.add(13);
            }
        } else if (i == 1) {
            if (count == 1) {
                if (svddk.size() > 24) {
                    
                        lstcountsv.add(13);
                        lstrow.add(svddk.size()-12);
                    } else {
                        lstcountsv.add(12);
                        lstrow.add(svddk.size()-11);
                    }
            } else {
                if (svddk.size() > 36) {
                    lstcountsv.add(12);
                    lstrow.add(14);
                } else {
                    lstcountsv.add(12);
                    lstrow.add(13);
                }
            }
        } else if (i == 2) {
            if (svddk.size() > 36) {
                lstcountsv.add(26);
                lstrow.add(svddk.size() - 24);
            } else if(svddk.size()>37){
                lstcountsv.add(25);
                lstrow.add(svddk.size() - 25);
            }else{
                lstcountsv.add(24);
                lstrow.add(svddk.size() - 23);
            }
        }

    }

    @Override
    public ArrayList<Sinhvien> docexceldiemdanh(Iterator<Row> iterquiz, List<Integer> lstcell) throws Exception {
        while (iterquiz.hasNext()) {
                Row row = iterquiz.next();
                Sinhvien mh1 = new Sinhvien();
                Iterator<Cell> celliter = row.cellIterator();
                if (row.getRowNum() == 2) {
                    //mamon = row.getCell(3).getStringCellValue();
                    lop = "com108.1";
                }
                if (row.getRowNum() == 3) {
                    //lop = row.getCell(3).getStringCellValue();
                    mamon=" (com108)";
                }
                if (row.getRowNum() == 4) {
                    mua = row.getCell(3).getStringCellValue();
                }
                if (row.getRowNum() > 7) {
                    while(celliter.hasNext()){
                        cell = celliter.next();
                        if (cell.getColumnIndex() == lstcell.get(2)) {
                            mh1.setTinhtrang(row.getCell(lstcell.get(2)).getStringCellValue());
                        }
                        if (cell.getColumnIndex() == lstcell.get(0)) {
                            mh1.setMasv(row.getCell(lstcell.get(0)).getStringCellValue());
                        }
                        if (cell.getColumnIndex() == lstcell.get(1)) {
                            mh1.setTensv(row.getCell(lstcell.get(1)).getStringCellValue());
                        }
                    }
                    mh1.setDiemonl(0);
                    lstsv.add(mh1);
                }
            }
            xuatkehoachthi(lop, mamonhoc(mamon));
        return this.lstsv;
    }

}
