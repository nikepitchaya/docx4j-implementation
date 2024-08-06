package com.example.demo;

import java.io.File;
import java.math.BigInteger;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.TcPrInner.GridSpan;
import org.docx4j.wml.CTTblPrBase.TblStyle;
import org.docx4j.wml.*;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.docx4j.jaxb.Context;

@SpringBootApplication
public class DemoApplication {
	public static void main(String[] args) {
		SpringApplication.run(DemoApplication.class, args);
	}

	@Bean
	public CommandLineRunner run() {
		return args -> {
			try {
				WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
				MainDocumentPart mainPart = wordPackage.getMainDocumentPart();
				String font = "TH Sarabun";
				mainPart.addObject(createTextStyle("รายงานการประชุมคณะอนุกรรมการพิจารณาอุทธรณ์เงินทดแทน", true, false,
						JcEnumeration.CENTER, 0, 0, font, 16));
				mainPart.addObject(createTextStyle("(กรณีเกี่ยวกับการแพทย์) (ชุดที่ 12)", true, false,
						JcEnumeration.CENTER, 0, 0, font, 16));
				mainPart.addObject(createTextStyle("ครั้งที่ 12/2566Type equation here.", true, false,
						JcEnumeration.CENTER, 0, 0, font, 16));
				mainPart.addObject(createTextStyle("(กรณีเกี่ยวกับการแพทย์) (ชุดที่ 12)", true, false,
						JcEnumeration.CENTER, 0, 0, font, 16));
				mainPart.addObject(createTextStyle("เมื่อวันพุธที่ 20 ธันวาคม 2566", true, false, JcEnumeration.CENTER,
						0, 0, font, 16));
				mainPart.addObject(createTextStyle("เมื่อวันพุธที่ 20 ธันวาคม 2566", true, false, JcEnumeration.CENTER,
						0, 0, font, 16));
				mainPart.addObject(createTextStyle("เวลา 13.30 น.", true, false, JcEnumeration.CENTER, 0, 0,
						font, 16));
				mainPart.addObject(
						createTextStyle("ณ ห้องประชุมสำนักงานกองทุนเงินทดแทน ชั้น 9 อาคารวิทุร แสงสิงแก้ว", true, false,
								JcEnumeration.CENTER, 0, 0, font, 16));
				mainPart.addObject(createTextStyle("และประชุมผ่านสื่ออิเล็กทรอนิกส์", true, false, JcEnumeration.CENTER,
						0, 0, font, 16));
				mainPart.addObject(createTextStyle("ผู้ประชุม", true, false, JcEnumeration.LEFT,
						0, 0, font, 16));
				mainPart.addObject(createTextStyle("1.\tศาสตราจารย์ภานุพันธ์\tทรงเจริญ\tประธานอนุกรรมการ", false, false,
						JcEnumeration.LEFT, 1000, 0, font, 16));
				mainPart.addObject(createTextStyle("2.\tรองศาสตราจารย์จุฑาไล\tตันฑเทอดธรรม\tอนุกรรมการ", false, false,
						JcEnumeration.LEFT, 1000, 0, font, 16));
				mainPart.addObject(createTextStyle("3.\tรองศาสตราจารย์เมธี\tวงศ์ศิริสุวรรณ\tอนุกรรมการ", false, false,
						JcEnumeration.LEFT, 1000, 0, font, 16));
				mainPart.addObject(createTextStyle("4.\tนายอำนาจ\tโชติชัย\tอนุกรรมการ", false, false,
						JcEnumeration.LEFT, 1000, 0, font, 16));
				// Create and add table
				mainPart.addObject(createTableStyle());

				mainPart.addObject(createTextStyle(
						"ประโยคความซ้อน (สังกรประโยค) หมายถึง ประโยคที่รวมประโยคความเดียว 1 ประโยคเป็นประโยคหลัก แล้วมีประโยคความเดียวอื่นมาเสริม มีข้อสังเกตคือ ประโยคหลัก (มุขยประโยค) กับ ประโยคย่อย (อนุประโยค) ของประโยคความช้อนมี น้ำหนักไม่เท่ากัน",
						true, false, JcEnumeration.LEFT,
						0, 0, font, 16));

				mainPart.addObject(createTextStyle(
						"ก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยวก๋วยเตี๋ยว",
						true, false, JcEnumeration.LEFT,
						0, 0, font, 16));
				File exportFile = new File("template.docx");
				wordPackage.save(exportFile);
			} catch (Exception e) {
				e.printStackTrace();
			}
		};
	}

	private Tbl createTableStyle() {
		// สร้างตาราง
		Tbl table = new Tbl();

		// ตั้งค่าตาราง
		TblPr tblPr = new TblPr();
		TblStyle tblStyle = new TblStyle();
		tblStyle.setVal("TableGrid");
		tblPr.setTblStyle(tblStyle);
		table.setTblPr(tblPr);

		// กำหนดคอลัมน์
		TblGrid tblGrid = new TblGrid();
		TblGridCol gridCol1 = new TblGridCol();
		gridCol1.setW(BigInteger.valueOf(2000));
		TblGridCol gridCol2 = new TblGridCol();
		gridCol2.setW(BigInteger.valueOf(2000));
		TblGridCol gridCol3 = new TblGridCol();
		gridCol3.setW(BigInteger.valueOf(2000));
		TblGridCol gridCol4 = new TblGridCol();
		gridCol4.setW(BigInteger.valueOf(2000));
		tblGrid.getGridCol().add(gridCol1);
		tblGrid.getGridCol().add(gridCol2);
		tblGrid.getGridCol().add(gridCol3);
		tblGrid.getGridCol().add(gridCol4);
		table.setTblGrid(tblGrid);

		Tr tableHeader1 = new Tr();
		Tr tableHeader2 = new Tr();

		Tc tc1 = new Tc();
		TcPr tcpr1 = new TcPr();
		TblWidth tbl1 = new TblWidth();
		tbl1.setW(BigInteger.valueOf(2000));
		tbl1.setType("dxa");
		P p1 = createTextStyle("ลำดับ", false, false, JcEnumeration.CENTER, 0, 0, "TH Sarabun", 16);
		tcpr1.setTcW(tbl1);
		tc1.getContent().add(p1);
		tc1.setTcPr(tcpr1);

		Tc tc2 = new Tc();
		TcPr tcpr2 = new TcPr();
		TblWidth tbl2 = new TblWidth();
		tbl2.setW(BigInteger.valueOf(2000));
		tbl2.setType("dxa");
		P p2 = createTextStyle("รายการ", false, false, JcEnumeration.CENTER, 0, 0, "TH Sarabun", 16);
		tcpr2.setTcW(tbl2);
		tc2.getContent().add(p2);
		tc2.setTcPr(tcpr2);

		Tc tc3 = new Tc();
		TcPr tcpr3 = new TcPr();
		TblWidth tbl3 = new TblWidth();
		tbl3.setW(BigInteger.valueOf(4000));
		tbl3.setType("dxa");
		P p3 = createTextStyle("รพ. การุญเวช ปทุมธานี", false, false, JcEnumeration.CENTER, 0, 0, "TH Sarabun", 16);
		tcpr3.setTcW(tbl3);
		tc3.getContent().add(p3);
		GridSpan gridSpan = new GridSpan();
		gridSpan.setVal(BigInteger.valueOf(2));
		tcpr3.setGridSpan(gridSpan);
		tc3.setTcPr(tcpr3);

		tableHeader1.getContent().add(tc1);
		tableHeader1.getContent().add(tc2);
		tableHeader1.getContent().add(tc3);
		table.getContent().add(tableHeader1);

		Tc tc4 = new Tc();
		TcPr tcpr4 = new TcPr();
		TblWidth tbl4 = new TblWidth();
		tbl4.setW(BigInteger.valueOf(4000));
		tbl4.setType("dxa");
		tcpr4.setTcW(tbl4);
		tc4.setTcPr(tcpr4);
		GridSpan gridSpan2 = new GridSpan();
		gridSpan2.setVal(BigInteger.valueOf(2));
		tcpr4.setGridSpan(gridSpan2);

		Tc tc5 = new Tc();
		TcPr tcpr5 = new TcPr();
		TblWidth tbl5 = new TblWidth();
		tbl5.setW(BigInteger.valueOf(2000));
		tbl5.setType("dxa");
		P p5 = createTextStyle("ผู้ป่วยใน\r\n" + //
				"14 – 18/12/65\r\n" + //
				"", false, false, JcEnumeration.CENTER, 0, 0, "TH Sarabun", 16);
		tcpr5.setTcW(tbl5);
		tc5.getContent().add(p5);
		tc5.setTcPr(tcpr5);

		Tc tc6 = new Tc();
		TcPr tcpr6 = new TcPr();
		TblWidth tbl6 = new TblWidth();
		tbl6.setW(BigInteger.valueOf(2000));
		tbl6.setType("dxa");
		P p6 = createTextStyle("ผู้ป่วยนอก\r\n" + //
				"14,19,20,30/12/65\r\n" + //
				"7,15,21/1/66\r\n" + //
				"18,28/2/66\r\n" + //
				"24/3/66\r\n" + //
				"", false, false, JcEnumeration.CENTER, 0, 0, "TH Sarabun", 16);
		tcpr6.setTcW(tbl6);
		tc6.getContent().add(p6);
		tc6.setTcPr(tcpr6);

		tableHeader2.getContent().add(tc4);
		tableHeader2.getContent().add(tc5);
		tableHeader2.getContent().add(tc6);
		table.getContent().add(tableHeader2);

		for (int i = 0; i < 20; i++) {
			Tr trData = new Tr();
			for (int j = 0; j < 4; j++) {
				Tc tcCell = new Tc();
				TcPr tcPrCell = new TcPr();
				TblWidth tblWidthCell = new TblWidth();
				tblWidthCell.setW(BigInteger.valueOf(2000));
				tblWidthCell.setType("dxa");
				tcPrCell.setTcW(tblWidthCell);
				P pdata = createTextStyle("ทดสอบ", false, false, JcEnumeration.CENTER, 0, 0, "TH Sarabun", 16);
				tcCell.getContent().add(pdata);
				tcCell.setTcPr(tcPrCell);
				trData.getContent().add(tcCell);
			}
			table.getContent().add(trData);
		}

		return table;
	}

	private P createTextStyle(String textValue, boolean isBold, boolean isItalic, JcEnumeration alignment,
			int marginLeft, int marginRight, String fontName, int fontSize) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		P paragraph = factory.createP();

		// Create and set paragraph properties
		PPr ppr = factory.createPPr();
		Jc justification = factory.createJc();
		justification.setVal(alignment);
		ppr.setJc(justification);

		// Set margins
		PPrBase.Ind indentation = factory.createPPrBaseInd();
		indentation.setLeft(BigInteger.valueOf(marginLeft));
		indentation.setRight(BigInteger.valueOf(marginRight));
		ppr.setInd(indentation);

		paragraph.setPPr(ppr);

		// Create and set run properties
		R run = factory.createR();
		Text text = factory.createText();
		text.setValue(textValue);
		run.getContent().add(text);

		RPr rpr = factory.createRPr();
		if (isBold) {
			BooleanDefaultTrue b = factory.createBooleanDefaultTrue();
			rpr.setB(b);
		}
		if (isItalic) {
			BooleanDefaultTrue i = factory.createBooleanDefaultTrue();
			rpr.setI(i);
		}

		RFonts rFonts = factory.createRFonts();
		rFonts.setAscii(fontName);
		rFonts.setHAnsi(fontName);
		rpr.setRFonts(rFonts);

		HpsMeasure size = factory.createHpsMeasure();
		size.setVal(BigInteger.valueOf(fontSize * 2)); // Font size in half-points
		rpr.setSz(size);
		rpr.setSzCs(size);

		run.setRPr(rpr);
		paragraph.getContent().add(run);

		return paragraph;
	}
}
