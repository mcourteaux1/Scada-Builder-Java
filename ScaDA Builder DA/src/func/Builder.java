/*
 * 
 */
package func;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * The Class Builder.
 */
@SuppressWarnings("static-access")
public class Builder {

	public static String state;
	public static String region;
	public static String aor;
	public static String scadar;
	public static String devType;
	public static String rtu;
	public static CellType rtuCell;
	public static CellType rtuAddCell;
	public static String rtuAdd;
	public static String ip;
	public static String dispVer;
	public static String lineKv;
	public static String menu;
	public static double devKv;

	public static void readEditSheet(String es) throws IOException {

		FileInputStream fis = new FileInputStream(new File(es));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		state = wb.getSheetAt(0).getRow(3).getCell(3).getStringCellValue();
		region = wb.getSheetAt(0).getRow(5).getCell(3).getStringCellValue();
		aor = wb.getSheetAt(0).getRow(9).getCell(3).getStringCellValue();
		scadar = wb.getSheetAt(0).getRow(3).getCell(7).getStringCellValue();
		devType = wb.getSheetAt(0).getRow(3).getCell(11).getStringCellValue();
		rtu = wb.getSheetAt(0).getRow(4).getCell(11).getRawValue().toString();
		rtuCell = wb.getSheetAt(0).getRow(4).getCell(11).getCellType();
		rtuAddCell = wb.getSheetAt(0).getRow(6).getCell(11).getCellType();
		rtuAdd = wb.getSheetAt(0).getRow(6).getCell(11).getRawValue().toString();
		ip = wb.getSheetAt(0).getRow(8).getCell(11).getStringCellValue();
		devKv = wb.getSheetAt(1).getRow(9).getCell(4).getNumericCellValue();
		wb.close();

		switch (devType) {

		case "IntelliRupter":
			devType = "IR";
			break;

		}

		if (rtuCell != rtuAddCell) {
			rtu = wb.getSheetAt(0).getRow(4).getCell(11).getStringCellValue();
		} else if (rtuCell == rtuAddCell) {
			rtu = wb.getSheetAt(0).getRow(4).getCell(11).getRawValue().toString();
		}

		scadar = wb.getSheetAt(0).getRow(3).getCell(7).getStringCellValue();

		if (scadar.contains("_D1_")) {
			dispVer = "D1";
		} else if (scadar.contains("_D2_")) {
			dispVer = "D2";
		} else if (scadar.contains("_D3_")) {
			dispVer = "D3";
		} else if (scadar.contains("_D4_")) {
			dispVer = "D4";
		} else if (scadar.contains("_D5_")) {
			dispVer = "D5";
		} else if (scadar.contains("_D6_")) {
			dispVer = "D6";
		} else if (scadar.contains("_D7_")) {
			dispVer = "D7";
		} else if (scadar.contains("_D8_")) {
			dispVer = "D8";
		} else if (scadar.contains("_D9_")) {
			dispVer = "D9";
		} else if (scadar.contains("_D10_")) {
			dispVer = "D10";
		} else if (scadar.contains("_D11_")) {
			dispVer = "D11";
		} else if (scadar.contains("_D12_")) {
			dispVer = "D12";
		} else if (scadar.contains("_D13_")) {
			dispVer = "D13";
		} else if (scadar.contains("_D14_")) {
			dispVer = "D14";
		} else if (scadar.contains("_D15_")) {
			dispVer = "D15";
		} else if (scadar.contains("_D16_")) {
			dispVer = "D16";
		} else if (scadar.contains("_D17_")) {
			dispVer = "D17";
		}

		switch (aor) {

		case "DOCNL":
			region = "ELLN";
			menu = "_NLA";
			break;
		case "DOCSL":
			region = "ELLS";
			menu = "_SLA";
			break;
		case "DOCSE":
			region = "ELLS";
			menu = "_SLA";
			break;
		case "DOCLR":
			region = "EAI";
			menu = "_AR";
			break;
		case "DOCMS":
			region = "EMI";
			menu = "_MPL";
			break;
		case "DOCMN":
			region = "EMI";
			menu = "_MPL";
			break;
		case "DOCBM":
			region = "ETI";
			menu = "_ETI";
			break;
		case "DOCNO":
			region = "ENOI";
			menu = "_SLA";
			break;
		case "DOCWL":
			region = "EGSL";
			menu = "_EGSL";
			break;
		case "DOCEL":
			region = "EGSL";
			menu = "_EGSL";
			break;
		}

		if (devKv <= 5) {
			lineKv = "_4KV";
		} else if (devKv > 15 && devKv < 30) {
			lineKv = "_25KV";
		} else if (devKv > 30) {
			lineKv = "_34KV";
		}

	}

	public String getState() {
		return this.state;
	}

	public String getRegion() {
		return this.region;
	}

	public String getAor() {
		return this.aor;
	}

	public String getScadar() {
		return this.scadar;
	}

	public String getDevType() {
		return this.devType;
	}

	public String getRtu() {
		return this.rtu;
	}

	public String getRtuAdd() {
		return this.rtuAdd;
	}

	public String getIp() {
		return this.ip;
	}

	public String getLineKv() {
		return this.lineKv;
	}

	public String getDispVer() {
		return this.dispVer;
	}

	public String getMenu() {
		return this.menu;
	}

	public CellType getRtuCell() {
		return this.rtuCell;
	}

	public CellType getRtuAddCell() {
		return this.rtuAddCell;
	}

	public static void displayGen(String rtu, String menu, String devType, String dispVer, String lineKv)
			throws IOException {

		PrintWriter myWriter = new PrintWriter("C:\\Users\\Mikey\\Desktop\\ScaDA Builder (Python)\\" + rtu + ".txt");

		myWriter.println("");
		myWriter.println("     display " + "\"" + rtu + "\"");
		myWriter.println("     (");
		myWriter.println(
				"         title(localize " + "\"" + "%DIS% [%DISAPP%][%DISFAM%][%HOST%]   (%VP%) %REF%" + "\"" + ")");
		myWriter.println("");
		myWriter.println("         application " + "\"" + "RECON" + "\"");
		myWriter.println("         (");
		myWriter.println("             color" + "(" + "\"" + "0,0,0" + "\"" + ")");
		myWriter.println("         )");
		myWriter.println("         application " + "\"" + "SCADA" + "\"");
		myWriter.println("         (");
		myWriter.println("             color" + "(" + "\"" + "0,0,0" + "\"" + ")");
		myWriter.println("         )");
		myWriter.println("         color" + "(" + "\"" + "0,0,0" + "\"" + ")");
		myWriter.println("         scale_to_fit_style(XY)");
		myWriter.println("         menu_bar_item " + "\"" + "SCADA_RELATED_DISPLAYS_MENU" + "\"" + "(");
		myWriter.println("         label(localize " + "\"" + "Related Displays" + "\"" + ")");
		myWriter.println("         set(" + "\"" + "ONELINES" + "\"" + ") )");
		myWriter.println("         menu_bar_item " + "\"" + "ONELINES" + menu + "\"" + "(");
		myWriter.println("         label(localize " + "\"" + "Onelines" + "\"" + ")");
		myWriter.println("         set(" + "\"" + "ONELINES_MENU" + "\"" + ") )");
		myWriter.println("         permitted_if");
		myWriter.println("         (");
		myWriter.println("             one_of(");
		myWriter.println("             class(");
		myWriter.println("             " + "\"" + "DSPTRWEA" + "\"" + ") )");
		myWriter.println("         )");
		myWriter.println("         horizontal_unit(10)");
		myWriter.println("         vertical_unit(10)");
		myWriter.println("         horizontal_page(50)");
		myWriter.println("         vertical_page(50)");
		myWriter.println("         refresh(4)");
		myWriter.println("         not locked_in_viewport");
		myWriter.println("         horizontal_scroll_bar");
		myWriter.println("         vertical_scroll_bar");
		myWriter.println("         std_menu_bar");
		myWriter.println("         not command_window");
		myWriter.println("         not on_top");
		myWriter.println("         not ret_last_tab_pnum");
		myWriter.println("         default_zoom(1.0000000)");
		myWriter.println("         simple_layer " + "\"" + "DEFAULT" + "\"" + "");
		myWriter.println("         (");
		myWriter.println("             not clip_to_regions");
		myWriter.println("             picture " + "\"" + "SCADA_BANNER_TO_TABULAR" + "\"" + "");
		myWriter.println("             (");
		myWriter.println("                 set(" + "\"" + "ONELINES" + "\"" + ")");
		myWriter.println("                 origin(0 0)");
		myWriter.println("                 xlocked");
		myWriter.println("                 ylocked");
		myWriter.println("             )");
		myWriter.println("             picture " + "\"" + "RTU_BANNER_RTUSTATE" + "\"" + "");
		myWriter.println("             (");
		myWriter.println("                 set(" + "\"" + "ONELINES" + "\"" + ")");
		myWriter.println("                 origin(978 2)");
		myWriter.println("                 xlocked");
		myWriter.println("                 ylocked");
		myWriter.println("             )");
		myWriter.println("             picture " + "\"" + "TO_RTU_8CHAR" + "\"" + "");
		myWriter.println("             (");
		myWriter.println("                 set(" + "\"" + "ONELINES" + "\"" + ")");
		myWriter.println("                 origin(994 8)");
		myWriter.println("                 xlocked");
		myWriter.println("                 ylocked");
		myWriter.println("                 composite_key");
		myWriter.println("                 (");
		myWriter.println("                     record(" + "\"" + "SUBSTN" + "\"" + ") record_key(" + "\"" + "COMMS"
				+ "\"" + ")");
		myWriter.println(
				"                     record(" + "\"" + "DEVTYP" + "\"" + ") record_key(" + "\"" + "RTU" + "\"" + ")");
		myWriter.println(
				"                     record(" + "\"" + "DEVICE" + "\"" + ") record_key(" + "\"" + rtu + "\"" + ")");
		myWriter.println(
				"                     record(" + "\"" + "POINT" + "\"" + ") record_key(" + "\"" + "STAT" + "\"" + ")");
		myWriter.println("                 )");
		myWriter.println("             )");
		myWriter.println("             picture " + "\"" + "DA_" + devType + "_" + dispVer + lineKv + "" + "\"" + "");
		myWriter.println("             (");
		myWriter.println("                 set(" + "\"" + "ONELINES_DA" + "\"" + ")");
		myWriter.println("                 origin(306 62)");
		myWriter.println("                 composite_key");
		myWriter.println("                 (");
		myWriter.println(
				"                     record(" + "\"" + "SUBSTN" + "\"" + ") record_key(" + "\"" + rtu + "\"" + ")");
		myWriter.println(
				"                     record(" + "\"" + "DEVTYP" + "\"" + ") record_key(" + "\"" + "RECL" + "\"" + ")");
		myWriter.println(
				"                     record(" + "\"" + "DEVICE" + "\"" + ") record_key(" + "\"" + rtu + "\"" + ")");
		myWriter.println("                     partial_key");
		myWriter.println("                 )");
		myWriter.println("             )");
		myWriter.println("             picture " + "\"" + "MAN_IN_STATION_DOC" + "\"" + "");
		myWriter.println("             (");
		myWriter.println("                 set(" + "\"" + "ONELINES" + "\"" + ")");
		myWriter.println("                 origin(340 0)");
		myWriter.println("                 xlocked");
		myWriter.println("                 ylocked");
		myWriter.println("                 composite_key");
		myWriter.println("                 (");
		myWriter.println(
				"                     record(" + "\"" + "SUBSTN" + "\"" + ") record_key(" + "\"" + rtu + "\"" + ")");
		myWriter.println(
				"                     record(" + "\"" + "DEVTYP" + "\"" + ") record_key(" + "\"" + "STN" + "\"" + ")");
		myWriter.println(
				"                     record(" + "\"" + "DEVICE" + "\"" + ") record_key(" + "\"" + "DOC" + "\"" + ")");
		myWriter.println(
				"                     record(" + "\"" + "POINT" + "\"" + ") record_key(" + "\"" + "MANS" + "\"" + ")");
		myWriter.println("                 )");
		myWriter.println("             )");
		myWriter.println("             text");
		myWriter.println("             (");
		myWriter.println("                 gab " + "\"" + "TEXT_TITLE" + "\"");
		myWriter.println("                 set(" + "\"" + "ONELINES" + "\"" + ")");
		myWriter.println("                 origin(524 5)");
		myWriter.println("                 xlocked");
		myWriter.println("                 ylocked");
		myWriter.println("                 localize " + "\"" + rtu + "\"" + "");
		myWriter.println("             )");
		myWriter.println("         )");
		myWriter.println("     );");
		myWriter.close();
	}

	@SuppressWarnings("resource")
	public static void alarm(String rtu) throws IOException {

		FileInputStream myxls = new FileInputStream(
				"C:\\Users\\Mikey\\Desktop\\ScaDA Builder Java\\mcourte\\Project Files\\DA\\AlarmLocation.xlsm");
		XSSFWorkbook studentsSheet = new XSSFWorkbook(myxls);
		XSSFSheet worksheet = studentsSheet.getSheetAt(0);
		int lastRow = worksheet.getLastRowNum();
		System.out.println(rtu);
		Row row = worksheet.createRow(++lastRow);
		System.out.println("Alarm row: " + lastRow);
		row.createCell(0).setCellValue("");
		row.createCell(1).setCellValue(rtu);
		row.createCell(9).setCellValue("display /app=scada/viewport=alarm_oneline %LOCID%");

		myxls.close();
		FileOutputStream output_file = new FileOutputStream(new File(
				"C:\\Users\\Mikey\\Desktop\\ScaDA Builder Java\\mcourte\\Project Files\\DA\\AlarmLocation.xlsm"));
		// write changes
		studentsSheet.write(output_file);
		output_file.close();
		System.out.println(" is successfully written");
	}

	@SuppressWarnings("resource")
	public static void fgdis(String rtu) throws IOException {

		FileInputStream myxls = new FileInputStream(
				"C:\\Users\\Mikey\\Desktop\\ScaDA Builder Java\\mcourte\\Project Files\\DA\\FullGraphicsDisplayRecords.xlsm");
		XSSFWorkbook studentsSheet = new XSSFWorkbook(myxls);
		XSSFSheet worksheet = studentsSheet.getSheetAt(0);
		int lastRow = worksheet.getLastRowNum();

		Row row = worksheet.createRow(++lastRow);
		System.out.println("fgdis last row: " + lastRow);
		row.createCell(0).setCellValue("");
		row.createCell(1).setCellValue(rtu);
		row.createCell(2).setCellValue(rtu);
		myxls.close();
		FileOutputStream output_file = new FileOutputStream(new File(
				"C:\\Users\\Mikey\\Desktop\\ScaDA Builder Java\\mcourte\\Project Files\\DA\\FullGraphicsDisplayRecords.xlsm"));
		// write changes
		studentsSheet.write(output_file);
		output_file.close();
		System.out.println(" is successfully written");

	}

	@SuppressWarnings("resource")
	public static void comm(String rtu, String aor, String feChoice, String chanChoice) throws IOException {

		FileInputStream myxls = new FileInputStream(
				"C:\\Users\\Mikey\\Desktop\\ScaDA Builder Java\\mcourte\\Project Files\\DA\\SCADA SCDA COMMS - Substation Hierarchy.xlsm");
		XSSFWorkbook commState = new XSSFWorkbook(myxls);
		XSSFSheet commandSheet = commState.getSheetAt(4);
		XSSFSheet discreteSheet = commState.getSheetAt(5);
		XSSFSheet genEquipSheet = commState.getSheetAt(7);

		int commandEnd = commandSheet.getLastRowNum();
		int discreteEnd = discreteSheet.getLastRowNum();
		int genEquipEnd = genEquipSheet.getLastRowNum();

		Row commandRow = commandSheet.createRow(++commandEnd);
		Row discreteRow = discreteSheet.createRow(++discreteEnd);
		Row genEquipRow = genEquipSheet.createRow(++genEquipEnd);

		// Create changes for the Command Sheet
		commandRow.createCell(1).setCellValue("GenericEquipment COMMS RTU_DA " + rtu + " STAT ENABLE");
		commandRow.createCell(4).setCellValue("ENABLE");
		commandRow.createCell(7).setCellValue("GenericEquipment COMMS RTU_DA " + rtu + " STAT");
		commandRow.createCell(8).setCellValue("ENABLE");
		commandRow.createCell(10).setCellValue("No");
		commandRow.createCell(11).setCellValue("Yes");
		commandRow.createCell(14).setCellValue("Yes");
		commandRow.createCell(15).setCellValue("Yes");
		commandRow.createCell(16).setCellValue("Yes");
		commandRow.createCell(17).setCellValue("No");
		commandRow.createCell(19).setCellValue("Yes");
		commandRow.createCell(22).setCellValue("ENABLE");
		commandRow.createCell(23).setCellValue("Closed");
		commandRow.createCell(24).setCellValue("No");
		commandRow.createCell(25).setCellValue("0");
		commandRow.createCell(27).setCellValue("No");
		commandRow.createCell(28).setCellValue("No");
		commandRow.createCell(29).setCellValue("No");
		commandRow.createCell(30).setCellValue("No");
		commandRow.createCell(31).setCellValue("No");
		commandRow.createCell(32).setCellValue("No");
		commandRow.createCell(33).setCellValue("LDAS" + feChoice);
		commandRow.createCell(35).setCellValue("Yes");
		commandRow.createCell(36).setCellValue("COMMS.RTU." + rtu + ".STAT.ENABLE");
		commandRow.createCell(37).setCellValue("JDAS" + feChoice);
		commandRow.createCell(39).setCellValue("10");
		commandRow.createCell(40).setCellValue("No");
		commandRow.createCell(42).setCellValue("0");
		commandRow.createCell(45).setCellValue("0");

		// Create changes for the Discrete Sheet
		discreteRow.createCell(1).setCellValue("GenericEquipment COMMS RTU_DA " + rtu + " STAT");
		discreteRow.createCell(2).setCellValue("Minute");
		discreteRow.createCell(5).setCellValue("No");
		discreteRow.createCell(6).setCellValue("Yes");
		discreteRow.createCell(7).setCellValue("Yes");
		discreteRow.createCell(8).setCellValue("No");
		discreteRow.createCell(9).setCellValue(aor);
		discreteRow.createCell(10).setCellValue("No");
		discreteRow.createCell(11).setCellValue("No");
		discreteRow.createCell(12).setCellValue("No");
		discreteRow.createCell(13).setCellValue("No");
		discreteRow.createCell(14).setCellValue("No");
		discreteRow.createCell(15).setCellValue("No");
		discreteRow.createCell(16).setCellValue("No");
		discreteRow.createCell(18).setCellValue("No");
		discreteRow.createCell(19).setCellValue("No");
		discreteRow.createCell(20).setCellValue("PSDO");
		discreteRow.createCell(22).setCellValue("RT");
		discreteRow.createCell(23).setCellValue("STAT");
		discreteRow.createCell(24).setCellValue("SMPF");
		discreteRow.createCell(25).setCellValue("Yes");
		discreteRow.createCell(26).setCellValue("STAT");
		discreteRow.createCell(31).setCellValue("No");
		discreteRow.createCell(32).setCellValue("No");
		discreteRow.createCell(33).setCellValue("No");
		discreteRow.createCell(34).setCellValue("No");
		discreteRow.createCell(35).setCellValue("No");
		discreteRow.createCell(36).setCellValue("No");
		discreteRow.createCell(37).setCellValue("No");
		discreteRow.createCell(38).setCellValue("No");
		discreteRow.createCell(39).setCellValue("No");
		discreteRow.createCell(40).setCellValue("No");
		discreteRow.createCell(41).setCellValue("No");
		discreteRow.createCell(42).setCellValue("No");
		discreteRow.createCell(43).setCellValue("No");
		discreteRow.createCell(44).setCellValue("No");
		discreteRow.createCell(45).setCellValue("No");
		discreteRow.createCell(47).setCellValue("Yes");
		discreteRow.createCell(48).setCellValue("No");
		discreteRow.createCell(49).setCellValue("No");
		discreteRow.createCell(50).setCellValue("0");
		discreteRow.createCell(51).setCellValue("999");
		discreteRow.createCell(52).setCellValue("RTU");
		discreteRow.createCell(53).setCellValue("SwitchPosition");
		discreteRow.createCell(55).setCellValue("-999");
		discreteRow.createCell(56).setCellValue("No");
		discreteRow.createCell(57).setCellValue("No");
		discreteRow.createCell(58).setCellValue("No");
		discreteRow.createCell(59).setCellValue("Telemetry");
		discreteRow.createCell(60).setCellValue("0");
		discreteRow.createCell(61).setCellValue("No");
		discreteRow.createCell(63).setCellValue("0");
		discreteRow.createCell(64).setCellValue("Minute");
		discreteRow.createCell(65).setCellValue("Yes");
		discreteRow.createCell(66).setCellValue("Yes");
		discreteRow.createCell(67).setCellValue("Yes");
		discreteRow.createCell(68).setCellValue("Yes");
		discreteRow.createCell(69).setCellValue("Yes");
		discreteRow.createCell(70).setCellValue("No");
		discreteRow.createCell(72).setCellValue("No");
		discreteRow.createCell(73).setCellValue("No");
		discreteRow.createCell(74).setCellValue("No");
		discreteRow.createCell(75).setCellValue("No");
		discreteRow.createCell(76).setCellValue("No");
		discreteRow.createCell(77).setCellValue("No");
		discreteRow.createCell(122).setCellValue("COMMS RTU " + rtu);
		discreteRow.createCell(157).setCellValue("No");
		discreteRow.createCell(158).setCellValue("No");
		discreteRow.createCell(159).setCellValue("Yes");
		discreteRow.createCell(160).setCellValue("Yes");
		discreteRow.createCell(165).setCellValue("COMMS.RTU." + rtu + ".STAT");
		discreteRow.createCell(166).setCellValue("PSDO");
		discreteRow.createCell(167).setCellValue("JDAS" + feChoice);
		discreteRow.createCell(170).setCellValue("No");
		discreteRow.createCell(171).setCellValue("LDAS" + feChoice);
		discreteRow.createCell(173).setCellValue("10");
		discreteRow.createCell(174).setCellValue("No");
		discreteRow.createCell(175).setCellValue("0");
		discreteRow.createCell(180).setCellValue("Yes");
		discreteRow.createCell(182).setCellValue("0");
		discreteRow.createCell(183).setCellValue("No");
		discreteRow.createCell(187).setCellValue("No");
		discreteRow.createCell(188).setCellValue("No");
		discreteRow.createCell(189).setCellValue("No");
		discreteRow.createCell(190).setCellValue("No");
		discreteRow.createCell(191).setCellValue("No");
		discreteRow.createCell(192).setCellValue("No");
		discreteRow.createCell(193).setCellValue("No");
		discreteRow.createCell(194).setCellValue("No");
		discreteRow.createCell(195).setCellValue("No");
		discreteRow.createCell(196).setCellValue("No");

		// Create changes for the Generic Equipment Sheet
		genEquipRow.createCell(1).setCellValue("COMMS RTU_DA " + rtu);
		genEquipRow.createCell(3).setCellValue("0");
		genEquipRow.createCell(8).setCellValue("No");
		genEquipRow.createCell(9).setCellValue("No");
		genEquipRow.createCell(12).setCellValue("No");
		genEquipRow.createCell(13).setCellValue("No");
		genEquipRow.createCell(15).setCellValue("0");
		genEquipRow.createCell(16).setCellValue("No");
		genEquipRow.createCell(18).setCellValue("No");
		genEquipRow.createCell(21).setCellValue("Yes");
		genEquipRow.createCell(22).setCellValue("No");
		genEquipRow.createCell(25).setCellValue(rtu);
		genEquipRow.createCell(35).setCellValue("RTU");
		genEquipRow.createCell(36).setCellValue("COMMS RTU");
		genEquipRow.createCell(38).setCellValue("No");
		genEquipRow.createCell(39).setCellValue("No");
		genEquipRow.createCell(40).setCellValue("No");
		genEquipRow.createCell(41).setCellValue("No");
		genEquipRow.createCell(42).setCellValue("No");
		genEquipRow.createCell(44).setCellValue("Yes");
		genEquipRow.createCell(45).setCellValue("No");
		genEquipRow.createCell(46).setCellValue("Yes");
		genEquipRow.createCell(47).setCellValue("No");
		genEquipRow.createCell(52).setCellValue("No");
		genEquipRow.createCell(53).setCellValue("No");
		genEquipRow.createCell(55).setCellValue("No");
		genEquipRow.createCell(57).setCellValue("No");
		genEquipRow.createCell(58).setCellValue("No");
		genEquipRow.createCell(59).setCellValue("No");
		genEquipRow.createCell(60).setCellValue("No");
		genEquipRow.createCell(61).setCellValue("No");
		genEquipRow.createCell(62).setCellValue("No");
		genEquipRow.createCell(65).setCellValue("No");
		genEquipRow.createCell(67).setCellValue(aor);
		genEquipRow.createCell(70).setCellValue("10");
		genEquipRow.createCell(72).setCellValue("COMMS.RTU." + rtu);
		genEquipRow.createCell(77).setCellValue("No");

		myxls.close();
		FileOutputStream output_file = new FileOutputStream(new File(
				"C:\\Users\\Mikey\\Desktop\\ScaDA Builder Java\\mcourte\\Project Files\\DA\\SCADA SCDA COMMS - Substation Hierarchy.xlsm"));
		// write changes
		commState.write(output_file);
		output_file.close();
		System.out.println(" is successfully written");

	}

	@SuppressWarnings("resource")
	public static void linker(String scadar, String rtu) throws IOException {

		FileInputStream dataItem = new FileInputStream(
				"C:\\Users\\Mikey\\Desktop\\ScaDA Builder Java\\mcourte\\Project Files\\DA.xlsm");
		XSSFWorkbook dataBook = new XSSFWorkbook(dataItem);
		XSSFSheet dataSheet = dataBook.getSheetAt(0);
		int dataEnd = dataSheet.getLastRowNum();
		int ctrlCol = 42;
		int anlgCol = 45;
		int sttsCol = 47;

		String anlg = " ANLG IED ";
		String ctrl = " CTRL IED ";
		String stts = " STTS IED ";
		String ge = "GenericEquipment ";
		String recl = " RECL ";

		// 351P
		String[] anlg351PD1 = { " AMPA", " AMPB", " AMPC", " AMPG", " AMW", " BMW", " CMW", " AMX", " BMX", " CMX",
				" VAZ", " VBZ", " VCZ", " VAY", " VBY", " VCY", " FIA", " FIB", " FIC", " PKUP", " VBAT", " WRA",
				" WRB", " WRC", " OPC3" };
		String[] ctrl351PD1 = { " STTS TRIP", " STTS CLOSE", " GRTR DISABLE", " GRTR ENABLE", " RCLS NONRECL",
				" RCLS AUTO", " ALST DISABLE", " ALST ENABLE", " SWMD ENABLE", " SWMD DISABLE", " TSTB OFF",
				" TSTB TEST" };
		String[] stts351PD1 = { " STTS", " GRTR", " RCLS", " RMLO", " TAPH", " TBPH", " TCPH", " TGND", " LOC", " RLYF",
				" TRBL", " BATB", " TSTB", " ACVL", " ALST", " SWMD", " CMLS" };

		// 351R
		String[] anlg351RD1 = { " AMPA", " AMPB", " AMPC", " AMPG", " AMW", " BMW", " CMW", " AMX", " BMX", " CMX",
				" VAZ", " VBZ", " VCZ", " VAY", " VBY", " VCY", " FIA", " FIB", " FIC", " PKUP", " VBAT", " WRA",
				" WRB", " WRC", " OPC3" };
		String[] ctrl351RD1 = { " STTS TRIP", " STTS CLOSE", " GRTR DISABLE", " GRTR ENABLE", " RCLS NONRECL",
				" RCLS AUTO", " ALST DISABLE", " ALST ENABLE", " SWMD ENABLE", " SWMD DISABLE", " TSTB OFF",
				" TSTB TEST" };
		String[] stts351RD1 = { " STTS", " GRTR", " RCLS", " RMLO", " BLOK", " TAPH", " TBPH", " TCPH", " TGND", " LOC",
				" RLYF", " TRBL", " BATB", " TSTB", " ACVL", " ALST", " SWMD", "CMLS" };

		// 351RD2
		String[] anlg351RD2 = { " AMPA", " AMPB", " AMPC", " AMPG", " FIA", " FIB", " FIC", " PKUP", " VBAT", " WRA",
				" WRB", " WRC", " OPC3" };
		String[] ctrl351RD2 = { " STTS TRIP", " STTS CLOSE", " GRTR DISABLE", " GRTR ENABLE", " RCLS NONRECL",
				" RCLS AUTO", " ALST DISABLE", " ALST ENABLE", " SWMD ENABLE", " SWMD DISABLE", " TSTB OFF",
				" TSTB TEST" };
		String[] stts351RD2 = { " STTS", " GRTR", " RCLS", " RMLO", " BLOK", " TAPH", " TBPH", " TCPH", " TGND", " LOC",
				" RLYF", " TRBL", " BATB", " TSTB", " ACVL", " ALST", " SWMD", "CMLS" };

		// 351RSD13
		String[] anlg351RSD13 = { " AMPX", " XMW", " XMX", " VXZ", " FIX", " PKUP", " VBAT", " WRX" };
		String[] ctrl351RSD13 = { " STTS TRIP", " STTS CLOSE", " RCLS NONRECL", " RCLS AUTO", " ALST DISABLE",
				" ALST ENABLE", " SWMD DISABLE", " SWMD ENABLE", " TSTB TEST", " TSTB OFF" };
		String[] stts351RSD13 = { " STTS", " RCLS", " RMLO", " BLOK", " TXPH", " LOC", " RLYF", " TRBL", " BATB",
				" TSTB", " ACVL", " ALST", " SWMD" };

		// 651R2D1
		String[] anlg651R2D1 = { " AMPA", " AMPB", " AMPC", " AMPG", " AMW", " BMW", " CMW", " AMX", " BMX", " CMX",
				" VAZ", " VBZ", " VCZ", " VAY", " VBY", " VCY", " FIA", " FIB", " FIC", " FIG", " PKUP", " VBAT",
				" WRA", " WRB", " WRC", " OPCA", " OPCB", " OPCC" };
		String[] ctrl651R2D1 = { " STTS TRIP", " STTS CLOSE", " GRTR DISABLE", " GRTR ENABLE", " RCLS NONRECL",
				" RCLS AUTO", " ALST DISABLE", " ALST ENABLE", " SWMD ENABLE", " SWMD DISABLE", " TSTB OFF",
				" TSTB TEST" };
		String[] stts651R2D1 = { " STTS", " STTA", " STTB", " STTC", " GRTR", " RCLS", " RMLO", " BLOK", " TAPH",
				" TBPH", " TCPH", " TGND", " LOCA", " LOCB", " LOCC", " RLYF", " TRBL", " BATB", " TSTB", " ACVL",
				" ALST", " SWMD" };

		// 651R2D1S11S11
		String[] anlg651R2D1S11 = { " AMPA", " AMPB", " AMPC", " AMPG", " AMW", " BMW", " CMW", " AMX", " BMX", " CMX",
				" VAZ", " VBZ", " VCZ", " VAY", " VBY", " VCY", " FIA", " FIB", " FIC", " PKUP", " VBAT", " WRA",
				" WRB", " WRC", " OPCA", " OPCB", " OPCC" };
		String[] ctrl651R2D1S11 = { " STTS TRIP", " STTS CLOSE", " GRTR DISABLE", " GRTR ENABLE", " RCLS NONRECL",
				" RCLS AUTO", " ALST DISABLE", " ALST ENABLE", " SWMD ENABLE", " SWMD DISABLE", " TSTB OFF",
				" TSTB TEST" };
		String[] stts651R2D1S11 = { " STTS", " STTA", " STTB", " STTC", " GRTR", " RCLS", " RMLO", " BLOK", " TAPH",
				" TBPH", " TCPH", " TGND", " LOCA", " LOCB", " LOCC", " RLYF", " TRBL", " BATB", " TSTB", " ACVL",
				" ALST", " SWMD" };
		// 651RAD4
		String[] anlg651RAD4 = { " AMPA", " AMPB", " AMPC", " AMPG", " FIA", " FIB", " FIC", " FIG", " PKUP", " VBAT",
				" WRA", " WRB", " WRC", " OPC3", };
		String[] ctrl651RAD4 = { " STTS TRIP", " STTS CLOSE", " GRTR DISABLE", " GRTR ENABLE", " RCLS NONRECL",
				" RCLS AUTO", " ALST DISABLE", " ALST ENABLE", " SWMD ENABLE", " SWMD DISABLE", " TSTB OFF",
				" TSTB TEST" };
		String[] stts651RAD4 = { " STTS", " GRTR", " RCLS", " RMLO", " BLOK", " TAPH", " TBPH", " TCPH", " TGND",
				" LOC", " RLYF", " TRBL", " BATB", " TSTB", " ACVL", " ALST", " SWMD" };

		// 651RAD4S11
		String[] anlg651RAD4S11 = { " AMPA", " AMPB", " AMPC", " AMPG", " FIA", " FIB", " FIC", " PKUP", " VBAT",
				" WRA", " WRB", " WRC", " OPC3", };
		String[] ctrl651RAD4S11 = { " STTS TRIP", " STTS CLOSE", " GRTR DISABLE", " GRTR ENABLE", " RCLS NONRECL",
				" RCLS AUTO", " ALST DISABLE", " ALST ENABLE", " SWMD ENABLE", " SWMD DISABLE", " TSTB OFF",
				" TSTB TEST" };
		String[] stts651RAD4S11 = { " STTS", " GRTR", " RCLS", " RMLO", " BLOK", " TAPH", " TBPH", " TCPH", " TGND",
				" LOC", " RLYF", " TRBL", " BATB", " TSTB", " ACVL", " ALST", " SWMD" };

		// IR20D2
		String[] anlgIR20D2 = { " AMPA", " AMPB", " AMPC", " AMPG", " AMW", " BMW", " CMW", " AMX", " BMX", " CMX",
				" VAX", " VBX", " VCX", " VAY", " VBY", " VCY", " FIA", " FIB", " FIC", " FIG", " PKUP", " VBAT",
				" WRA", " WRB", " WRC", " OPCA", " OPCB", " OPCC" };
		String[] ctrlIR20D2 = { " STTS TRIP", " STTS CLOSE", " GRTP ENABLE", " GRTP DISABLE", " RCBL AUTO",
				" RCBL NONRECL", " ALTR ON", " ALTR OFF", " TSTB TEST", " TSTB OFF", " HRDC DISABLE", " HRDC ENABLE",
				" CMOP RESET", " CMOP ALARM" };
		String[] sttsIR20D2 = { " STTS", " STTA", " STTB", " STTC", " GRTR", " RCLS", " RMLO", " BLOK", " TAPH",
				" TBPH", " TCPH", " TGND", " LOCA", " LOCB", " LOCC", " RLYF", " TRBL", " BATB", " TSTB", " ACVL",
				" TMNR", " HRDC", "CMOP", "HOTL", "RECF" };

		// IR20D17
		String[] anlgIR20D17 = { " AMPA", " AMPB", " AMPC", " AMPG", " AMW", " BMW", " CMW", " AMX", " BMX", " CMX",
				" VAX", " VBX", " VCX", " VAY", " VBY", " VCY", " FIA", " FIB", " FIC", " FIG", " PKUP", " VBAT",
				" WRA", " WRB", " WRC", " OPCA", " OPCB", " OPCC" };
		String[] ctrlIR20D17 = { " STTS TRIP", " STTS CLOSE", " GRTP ENABLE", " GRTP DISABLE", " RCBL AUTO",
				" RCBL NONRECL", " TSTB OFF", " TSTB TEST", " HRDC ENABLE", " HRDC DISABLE" };
		String[] sttsIR20D17 = { " STTS", " STTA", " STTB", " STTC", " GRTR", " RCLS", " RMLO", " BLOK", " TAPH",
				" TBPH", " TCPH", " TGND", " LOCA", " LOCB", " LOCC", " RLYF", " TRBL", " BATB", " TSTB", " ACVL",
				" ALST", " SWMD" };
		System.out.println(dataEnd);

		switch (scadar) {
		case "RECL_351P VIA RTAC_30_S5_HE7_D1":
			for (int i = 1; i <= dataEnd; i++) {
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0000")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0001")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0002")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0003")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0004")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0005")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0006")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0007")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0008")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0009")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0010")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0011")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0012")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0013")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0014")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0015")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0016")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0017")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[17]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0018")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[18]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0019")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[19]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0020")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[20]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0021")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[21]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0022")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[22]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0023")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351PD1[23]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 1")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351PD1[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 2")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351PD1[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 3")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351PD1[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 4")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351PD1[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 5")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351PD1[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 6")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351PD1[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 7")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351PD1[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 8")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351PD1[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 10")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351PD1[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 9")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351PD1[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 11")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351PD1[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 12")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351PD1[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0000")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0001")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0002")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0003")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0004")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0005")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0006")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0007")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0008")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0009")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0010")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0011")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0012")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0013")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0014")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0015")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0016")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0017")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351PD1[17]);
					break;
				}

			}

		case "RECL_351R VIA RTAC_30_S5_HE8_D1":
			for (int i = 1; i <= dataEnd; i++) {
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0000")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0001")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0002")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0003")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0004")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0005")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0006")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0007")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0008")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0009")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0010")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0011")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0012")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0013")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0014")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0015")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0016")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0017")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[17]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0018")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[18]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0019")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[19]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0020")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[20]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0021")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[21]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0022")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[22]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0023")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD1[23]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 1")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD1[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 2")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD1[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 3")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD1[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 4")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD1[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 5")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD1[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 6")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD1[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 7")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD1[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 8")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD1[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 10")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD1[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 9")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD1[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 11")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD1[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 12")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD1[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0000")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0001")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0002")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0003")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0004")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0005")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0006")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0007")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0008")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0009")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0010")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0011")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0012")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0013")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0014")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0015")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0016")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0017")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[17]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0018")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD1[18]);
					break;
				}

			}

		case "RECL_351R VIA RTAC_30_S5_HE9_D2":
			for (int i = 1; i <= dataEnd; i++) {
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0000")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0001")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0002")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0003")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0004")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0005")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0006")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0007")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0008")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0009")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0010")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0011")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0012")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0013")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0014")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RD2[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 1")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD2[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 2")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD2[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 3")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD2[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 4")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD2[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 5")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD2[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 6")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD2[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 7")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD2[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 8")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD2[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 10")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD2[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 9")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD2[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 11")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD2[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 12")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RD2[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0000")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0001")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0002")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0003")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0004")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0005")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0006")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0007")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0008")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0009")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0010")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0011")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0012")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0013")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0014")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0015")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0016")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0017")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[17]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0018")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RD2[18]);
					break;
				}

			}

		case "RECL_351RS_30_S1_HE10_D13_EAY":
			for (int i = 1; i <= dataEnd; i++) {
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0000")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RSD13[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0001")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RSD13[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0002")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RSD13[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0003")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RSD13[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0004")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RSD13[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0005")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RSD13[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0006")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RSD13[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0007")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RSD13[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0008")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg351RSD13[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 1")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RSD13[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 2")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RSD13[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 3")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RSD13[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 4")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RSD13[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 5")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RSD13[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 6")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RSD13[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 7")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RSD13[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 8")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RSD13[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 10")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RSD13[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 9")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl351RSD13[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0000")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0001")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0002")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0003")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0004")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0005")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0006")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0007")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0008")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0009")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0010")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0011")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0012")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0013")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts351RSD13[13]);
					break;
				}
			}

		case "RECL_651R2_30_S1_HE1_D1_EAY":
			for (int i = 1; i <= dataEnd; i++) {
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0000")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0001")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0002")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0003")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0004")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0005")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0006")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0007")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0008")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0009")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0010")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0011")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0012")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0013")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0014")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0015")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0016")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0017")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[17]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0018")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[18]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0019")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[19]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0020")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[20]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0021")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[21]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0022")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[22]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0023")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[23]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0024")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[24]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0025")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[25]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0026")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[26]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0027")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1[27]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 1")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 2")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 3")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 4")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 5")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 6")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 7")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 8")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 10")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 9")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 11")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 12")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0000")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0001")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0002")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0003")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0004")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0005")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0006")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0007")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0008")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0009")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0010")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0011")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0012")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0013")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0014")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0015")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0016")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0017")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[17]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0018")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[18]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0019")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[19]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0020")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[20]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0021")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1[21]);
					break;
				}

			}
		case "RECL_651R2_30_S1-1_HE1-1_D1_EAY":
			for (int i = 1; i <= dataEnd; i++) {
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0000")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0001")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0002")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0003")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0004")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0005")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0006")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0007")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0008")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0009")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0010")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0011")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0012")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0013")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0014")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0015")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0016")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0017")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[17]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0018")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[18]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0019")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[19]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0020")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[20]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0021")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[21]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0022")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[22]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0023")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[23]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0024")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[24]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0025")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[25]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0026")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651R2D1S11[26]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 1")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1S11[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 2")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1S11[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 3")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1S11[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 4")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1S11[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 5")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1S11[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 6")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1S11[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 7")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1S11[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 8")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1S11[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 10")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1S11[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 9")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1S11[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 11")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1S11[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 12")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651R2D1S11[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0000")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0001")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0002")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0003")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0004")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0005")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0006")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0007")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0008")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0009")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0010")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0011")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0012")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0013")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0014")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0015")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0016")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0017")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[17]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0018")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[18]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0019")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[19]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0020")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[20]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0021")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651R2D1S11[21]);
					break;
				}

			}

		case "RECL_651RA_30_S1_HE4_D4_EAY":
			for (int i = 1; i <= dataEnd; i++) {
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0000")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0001")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0002")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0003")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0004")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0005")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0006")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0007")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0008")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0009")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0010")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0011")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0012")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0013")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0014")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 1")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 2")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 3")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 4")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 5")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 6")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 7")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 8")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 10")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 9")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 11")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 12")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0000")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0001")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0002")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0003")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0004")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0005")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0006")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0007")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0008")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0009")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0010")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0011")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0012")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0013")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0014")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0015")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0016")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0017")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4[17]);
					break;
				}

			}

		case "RECL_651RA_30_S1-1_HE4-1_D4_EAY":
			for (int i = 1; i <= dataEnd; i++) {
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0000")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0001")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0002")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0003")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0004")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0005")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0006")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0007")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0008")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0009")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0010")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0011")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0012")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0013")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlg651RAD4S11[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 1")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4S11[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 2")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4S11[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 3")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4S11[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 4")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4S11[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 5")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4S11[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 6")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4S11[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 7")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4S11[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 8")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4S11[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 10")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4S11[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 9")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4S11[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 11")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4S11[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 12")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrl651RAD4S11[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0000")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0001")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0002")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0003")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0004")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0005")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0006")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0007")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0008")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0009")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0010")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0011")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0012")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0013")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0014")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0015")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0016")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0017")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + stts651RAD4S11[17]);
					break;
				}

			}

		case "RECL_IR_20_S2_HE2_D2_EAY":
			for (int i = 1; i <= dataEnd; i++) {
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0000")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0001")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0002")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0003")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0004")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0005")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0006")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0007")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0008")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0009")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0010")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0011")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0012")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0013")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0014")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0015")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0016")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0017")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[17]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0018")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[18]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0019")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[19]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0020")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[20]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0021")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[21]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + anlg + "0022")) {
					dataSheet.getRow(i).createCell(anlgCol).setCellValue(ge + rtu + recl + rtu + anlgIR20D2[22]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 1")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrlIR20D2[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0000 2")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrlIR20D2[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 3")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrlIR20D2[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0001 4")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrlIR20D2[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 5")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrlIR20D2[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0002 6")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrlIR20D2[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 7")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrlIR20D2[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0003 8")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrlIR20D2[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 10")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrlIR20D2[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0004 9")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrlIR20D2[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 11")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrlIR20D2[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + ctrl + "0005 12")) {
					dataSheet.getRow(i).createCell(ctrlCol).setCellValue(ge + rtu + recl + rtu + ctrlIR20D2[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0000")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[0]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0001")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[1]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0002")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[2]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0003")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[3]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0004")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[4]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0005")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[5]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0006")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[6]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0007")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[7]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0008")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[8]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0009")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[9]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0010")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[10]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0011")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[11]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0012")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[12]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0013")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[13]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0014")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[14]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0015")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[15]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0016")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[16]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0017")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[17]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0018")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[18]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0019")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[19]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0020")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[20]);
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0021")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[21]);
					break;
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0022")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[22]);
					break;
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0023")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[23]);
					break;
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0024")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[24]);
					break;
				}
				if (dataSheet.getRow(i).getCell(1).getStringCellValue().equals(rtu + stts + "0025")) {
					dataSheet.getRow(i).createCell(sttsCol).setCellValue(ge + rtu + recl + rtu + sttsIR20D2[25]);
					break;
				}

			}

		case "IR20D17":

			break;

		case "TBD1":

			break;

		case "TBD2":

			break;

		case "TBD3":

			break;

		}

		dataItem.close();
		FileOutputStream output_file = new FileOutputStream(
				new File("C:\\Users\\Mikey\\Desktop\\ScaDA Builder Java\\mcourte\\Project Files\\DA\\DataItem.xlsm"));

		dataBook.write(output_file);
		output_file.close();
		System.out.println(" is successfully written");

	}

}
