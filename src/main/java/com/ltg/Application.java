package com.ltg;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Application {
	
	static String inRegisterFile;
	static String inReserveFile;
	static String outResultFile;
	static String inFilePath;

	public static void main(String[] args) {
//		final String inRegisterFile = args[0];
//		final String inReserveFile = args[1];
		
//		String inRegisterFile = "C:\\Users\\emylyn.audemard\\Documents\\CrossReference\\input\\HHRawSignUpData_ABIitsSubsidiariesandAffiliates_validated.xlsx";
//		String inHHFile = "C:\\Users\\emylyn.audemard\\Documents\\CrossReference\\input\\Asia Brewery, Inc. (ABI) and Subsidiaries Daily Report_reservation_2021-04-14_(01_30_00)_FINAL.xlsx";
		String outfolderPath = "C:\\Users\\emylyn.audemard\\Documents\\CrossReference\\err";
		
		String inFile = "C:\\Users\\emylyn.audemard\\Documents\\CrossReference\\Input";
		
		setInFilePath(inFile);
		
		File directoryPath = new File(getInFilePath());
		String excelFile[] = directoryPath.list();
		System.out.println("File Set-up.......");
		for(int i=0; i<excelFile.length; i++) {
			
			
			System.out.println();
			
			if(excelFile[i].contains("Daily")) {
				setInReserveFile(inFile+"\\"+excelFile[i]);
				
				setOutResultFile(outfolderPath+"\\"+getFileNameResult(getInReserveFile()));
				
			}
			
			if(excelFile[i].contains("HHRaw")) {
				setInRegisterFile(inFile+"\\"+excelFile[i]);
			}
		}
		
		System.out.println(getInRegisterFile());
		System.out.println(getInReserveFile());
		
		System.out.println("");
		
		if(!checkFile(".xlsx", getInRegisterFile()) || !checkFile(".xlsx", getInReserveFile())) {
			System.exit(0);
		}else {
			System.out.println("files are Valid.");
			System.out.println("");
		}
		
		//process files
//		List<HashMap<String, Object>> arrListFromRegFile = getDataFromRegisterFile(getInRegisterFile());
		List<HashMap<String, Object>> arrListFromRegFile = getValidControlNumberFormat(getDataFromRegisterFile(getInRegisterFile()));
		List<HashMap<String, Object>> arrListFromReserveFile = getDataFromReservationFile(getInReserveFile());
		
		
		for(HashMap<String, Object> s : arrListFromRegFile) {
//			System.out.println(s);
		}
		
		for(HashMap<String, Object> s : arrListFromReserveFile) {
//			System.out.println(s);
		}
		
		System.out.println("");System.out.println("");
		
		ArrayList<ArrayList<String>> errorList = new ArrayList<ArrayList<String>>();
		ArrayList<ArrayList<String>> errorListFormat = new ArrayList<ArrayList<String>>();
		
		List<String> ctrlNumberListReg = convertToList(arrListFromRegFile);
		
//		errorList.add(excessCtrlNUmber(ctrlNumberListReg, arrListFromReserveFile));
		errorList.add(excessCtrlNUmberv2(ctrlNumberListReg, arrListFromReserveFile));
		errorList.add(checkNoHHReg(ctrlNumberListReg, arrListFromReserveFile));
		errorList.add(findDuplicateInRegisterFile(arrListFromRegFile));
		errorList.add(findDuplicateInHHFile(arrListFromReserveFile));
		
		errorListFormat.add(checkInvalidControlNumberRegFile(arrListFromRegFile));
		
//		
//		if(checkInvalidControlNumberRegFile(arrListFromRegFile).size() != 0) {
//			System.out.println("Error in Control number format");
//			writeTextLogFile(errorListFormat);
//		}else {
//			writeTextLogFile(errorList);
//		}
		writeTextLogFile(errorList);
	}

	private static ArrayList<String> excessCtrlNUmberv2(List<String> ctrlNumberListReg,
			List<HashMap<String, Object>> arrListFromReserveFile) {
		
		
		int modCounter = 0;
		int covCounter = 0;
		ArrayList<String> errList = new ArrayList<String>();
		List<Object> EmpIdNotExist = new ArrayList<Object>();
		
		//create map <ctrlnumber, empnumber> from ctrlNumberListReg <ctrlnumber>
		HashMap<String, Object> mapListReg = convertListToHashMap(ctrlNumberListReg);
		
		//create list of employeeNumber from mapListReg
		List<Object> listReg = new ArrayList<Object>(mapListReg.values());
		
		// remove duplicate from listReg
		List<Object> listRegUnique = listReg.stream()
					     .distinct()
					     .collect(Collectors.toList());
		
//		System.out.println(mapListReg.size());
		
		for(Object r : listRegUnique) {

			for(String s : ctrlNumberListReg) {
				String[] empNum = s.toString().trim().split("_");
				if(r.toString().equals(empNum[1])) {
					
					if(empNum[2].contains("M")) {
						modCounter++;
					} else if(empNum[2].contains("C")) {
						covCounter++;
					}
				}
				
			}
//			System.out.println(r.toString()+" = M "+modCounter+"  C "+covCounter);
			
			for(HashMap<String, Object> m : arrListFromReserveFile) {
				if(r.toString().equals(m.get("employeeNumber").toString())) {
//					System.out.println("yes "+r.toString());
//					System.out.println("Employee Number "+ m.get("employeeNumber") +
//					": Moderna Control Number. "+
//					modCounter+" in Registration - "+
//					m.get("modernaOrders").toString()+" in Reservation. ");   
					
					//Incomplete
					if(Integer.valueOf((String) m.get("modernaOrders")) > modCounter) {
						
						errList.add("Error: "+m.get("firstName")+" "+m.get("lastName")+" with Employee Number "+ m.get("employeeNumber") +
								": Incomplete registration for Moderna. "+
								modCounter+" in Registration - "+
								m.get("modernaOrders").toString()+" in Reservation. ");   
					}
					
					if(Integer.valueOf((String) m.get("covovaxOrders")) > covCounter) {
						System.out.println("matched covovaxOrders");
						
						errList.add("Error: "+m.get("firstName")+" "+m.get("lastName")+" with Employee Number "+ m.get("employeeNumber") +
								": Incomplete registration for Covovax. "+
								covCounter+" in Registration - "+
								m.get("covovaxOrders").toString()+" in Reservation. ");   
					}
					
					//Excess Registration
					if(Integer.valueOf((String) m.get("modernaOrders")) < modCounter) {
						errList.add("Error: "+m.get("firstName")+" "+m.get("lastName")+" with Employee Number "+ m.get("employeeNumber") +
								": Excess registration for Moderna. "+
								modCounter+" in Registration - "+
								m.get("modernaOrders").toString()+" in Reservation. ");  
					}
					
					if(Integer.valueOf((String) m.get("covovaxOrders")) < covCounter) {
						errList.add("Error: "+m.get("firstName")+" "+m.get("lastName")+" with Employee Number "+ m.get("employeeNumber") +
								": Excess registration for Covovax. "+
								modCounter+" in Registration - "+
								m.get("covovaxOrders").toString()+" in Reservation. ");  
					}
				}
			}
			modCounter = 0;
			covCounter = 0;
		}
		
//		// remove duplicate from listReg
//		List<Object> EmpIdNotExistUnique = EmpIdNotExist.stream()
//					     .distinct()
//					     .collect(Collectors.toList());
//		
//		
//		//show Employee ID does not exist
//		for(Object s : EmpIdNotExistUnique) {
//			System.out.println("no "+s);
//		}
		
		return errList;
	}

	private static List<HashMap<String, Object>> getValidControlNumberFormat(List<HashMap<String, Object>> arrListFromRegFile) {
		List<HashMap<String, Object>> listsMap = new ArrayList<HashMap<String, Object>>();
		
		for(HashMap<String, Object> s : arrListFromRegFile) {
			String regex = "\\b"+s.get("companyCode")+"_[A-Za-z0-9-]+[_]{1}[M|C]{1}[1-9]{1}[0-9]?$";
			String regexPalex = "\\bPALEX_[A-Za-z0-9-]+[_]{1}[M|C]{1}[1-9]{1}[0-9]?$";
			
			String[] ctrlNumberArr = s.get("controlNumber").toString().replaceAll(" ", "").trim().split(",");
			boolean isValid;
			
			if(ctrlNumberArr.length > 1) {
				s.put("isRed", true);
			}else {
				isValid = Pattern.matches(regex, s.get("controlNumber").toString().trim());
				
				if(s.get("companyCode").toString() == "PAL") {
					if(!isValid) {
						isValid = Pattern.matches(regexPalex, s.get("controlNumber").toString());
						
						if(!isValid) {
							s.put("isRed", true);
						} else {
							s.put("isRed", false);
						}
					}else {
						s.put("isRed", false);
					}
				}else {
					if(!isValid) {
						s.put("isRed", true);
					} else {
						s.put("isRed", false);
					}
				}
			}
			listsMap.add(s);
		}
		
//		for(HashMap<String, Object> s : listsMap) {
//			System.out.println(s);
////			
//			if(Boolean.valueOf(s.get("isRed").toString()) == true) {
//				System.out.println(s.get("isRed")+" -- "+ s.get("controlNumber"));
//			}
//		}
////		
		
		return arrListFromRegFile;
	}

	private static ArrayList<String> checkInvalidControlNumberRegFile(List<HashMap<String, Object>> arrListFromRegFile) {
		
		ArrayList<String> errList = new ArrayList<String>();
		
		for(HashMap<String, Object> r : arrListFromRegFile) {
			String name = r.get("firstName")+" "+r.get("lastName");
			if(Boolean.valueOf((boolean) r.get("isRed"))) {
				errList.add("Error: Row "+r.get("regRowNumber")+" "+ name+" : Invalid Control Number "+r.get("controlNumber"));
			}
		}
		
		return errList;
	}

	private static void writeTextLogFile(ArrayList<ArrayList<String>> errorList) {
		try {
			FileWriter writer = new FileWriter(getOutResultFile(), true);
			BufferedWriter bufferedWriter = new BufferedWriter(writer);
		        
			for(ArrayList<String> al : errorList) {
				for(String r : al) {
					System.out.println(r);
					bufferedWriter.write(r);
					bufferedWriter.newLine();
				}
			}
			
			bufferedWriter.close();
			System.out.println("written successfully");
			System.out.println("Output File Location: "+getOutResultFile());
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
      
	}

	private static ArrayList<String> findDuplicateInHHFile(List<HashMap<String, Object>> arrListFromReserveFile) {
		
		ArrayList<String> ctrlNumberList = new ArrayList<String>();
		ArrayList<String> cnl = new ArrayList<String>();
		for(HashMap<String, Object> r : arrListFromReserveFile) {
			
			String[] Mctrln = r.get("ModernactrlNumber").toString().replaceAll(" ", "").trim().split(",");
			String[] Cctrln = r.get("CovovaxctrlNumber").toString().replaceAll(" ", "").trim().split(",");
			
			if(r.get("ModernactrlNumber").toString() != "--") {
				for(String s : Mctrln) {
					ctrlNumberList.add(s);
				}
			}
			
			if(r.get("CovovaxctrlNumber").toString() != "--") {
				for(String s : Cctrln) {
					ctrlNumberList.add(s);
				}
			}
		 
			
		}
		
		ArrayList<String> errList = new ArrayList<String>();
		List<String> seg = new ArrayList<String>();
		int counter = 0;
		Set<String> dupes = new HashSet<String>();
        for (String i : ctrlNumberList) {
            if (!dupes.add(i)) {
				seg.add(i);
            }
        }
//        System.out.println(seg);
        
        Set<String> uniqueSet = new HashSet<>(seg);
        List<String> nowList = new ArrayList<>(uniqueSet);
        
        List<String> rowsWithDupModerna = new ArrayList<String>();
        List<String> rowsWithDupcovovax = new ArrayList<String>();
        
        for(String se : nowList) {
        	for(HashMap<String, Object> r : arrListFromReserveFile) {
        		String[] Mctrln = r.get("ModernactrlNumber").toString().replaceAll(" ", "").trim().split(",");
    			String[] Cctrln = r.get("CovovaxctrlNumber").toString().replaceAll(" ", "").trim().split(",");
    			
    			for(String m : Mctrln) {
    				if( m.equals(se) ) {
    					rowsWithDupModerna.add(r.get("resRowNumber").toString());
    				}
    			}
    			
    			for(String c : Cctrln) {
    				if( c.equals(se) ) {
    					rowsWithDupcovovax.add(r.get("resRowNumber").toString());
    				}
    			}
        	}
        	
        	String[] ctrlNum = se.toString().replace(" ", "").trim().split("_");
        	if(ctrlNum[2].contains("M")) {
             	errList.add("Error: Row "+rowsWithDupModerna+" Control Number : "+se+"  Found duplicate in Household Reservation File");
        	}else if(ctrlNum[2].contains("C")) {
             	errList.add("Error: Row "+rowsWithDupcovovax+" Control Number : "+se+"  Found duplicate in Household Reservation File");
        	}

     	
     	rowsWithDupModerna.clear();
     	rowsWithDupcovovax.clear();


        }
        
        return errList;
		
	}

	private static ArrayList<String> findDuplicateInRegisterFile(List<HashMap<String, Object>> arrListFromRegFile) {
		
		//todo - Error: Row[1,2,3] Control Number : APL_923_M2  Found duplicate in Registration File
		List<String> rowsWithDup = new ArrayList<String>();
		List<String> ctrlNumberList = convertToList(arrListFromRegFile);
		ArrayList<String> errList = new ArrayList<String>();
		List<String> seg = new ArrayList<String>();
		int counter = 0;
		Set<String> dupes = new HashSet<String>();
        for (String i : ctrlNumberList) {
            if (!dupes.add(i)) {
				seg.add(i);
            }
        }
        
//        System.out.println(ctrlNumberList);
        
        Set<String> uniqueSet = new HashSet<>(seg);
        List<String> nowList = new ArrayList<>(uniqueSet);
        
        for(String se : nowList) {
         	
         
        	for(HashMap<String, Object> s : arrListFromRegFile) {
    			System.out.println(s);
    			
    			if(se.equals(s.get("controlNumber").toString().replaceAll(" ", "").trim())){
    				
					rowsWithDup.add(s.get("regRowNumber").toString());
    			}
    		}
        	errList.add("Error: Row "+rowsWithDup+" Control Number : "+se+"  Found duplicate in Registration File");
        	rowsWithDup.clear();
//        	for(HashMap<String, Object> r : arrListFromRegFile) {
//        		if(se.equals(r.get("controlNumber").toString())) {
//					counter++;
//        		}
//        	}
////        	System.out.println("Control Number "+se+" in Registration File found "+counter+" Duplicates");
//        	errList.add("Control Number "+se+" in Registration File found "+counter+" Duplicates");
//        	counter = 0;
//        	
//        	for(HashMap<String, Object> r : arrListFromRegFile) {
//        		if(se.equals(r.get("controlNumber").toString())) {
////					System.out.println("----"+r.get("firstName")+" "+r.get("lastName")+" with Control Number "+ r.get("controlNumber") +
////							": Duplicate Control Number. ");
//					errList.add("----"+r.get("firstName")+" "+r.get("lastName")+" with Control Number "+ r.get("controlNumber") +
//							": Duplicate Control Number. ");
//        		}
//        	}
        }
        
        return errList;
	}

	private static ArrayList<String> checkNoHHReg(List<String> arrListFromRegFile,
			List<HashMap<String, Object>> arrListFromReserveFile) {
		
		for(HashMap<String, Object> s : arrListFromReserveFile) {
//			System.out.println(s);
		}
		
		ArrayList<String> errList = new ArrayList<String>();
		
		HashMap<String, List<String>> mapArrListReg =  new HashMap<>();
		HashMap<String, List<String>> mapArrListRes =  new HashMap<>();
		
		mapArrListReg = groupByEmployeeNumber(convertListToHashMap(arrListFromRegFile));
		mapArrListRes = groupByEmployeeNumber(convertListHashToHashMap(arrListFromReserveFile));
		
		HashSet<String> unionKeys = new HashSet<>(mapArrListReg.keySet());
		unionKeys.addAll(mapArrListRes.keySet());
		 
		unionKeys.removeAll(mapArrListReg.keySet());
		 
		for(String arr : unionKeys) {
//			System.out.println(arr);
			
			for(HashMap<String, Object> s : arrListFromReserveFile) {
				
//				System.out.println(s);
				if(s.containsValue(arr)) {
//					System.out.println(s.get("firstName")+" "+s.get("lastName")+" with Employee Number "+ s.get("employeeNumber") +
//							": No HH that registered. ");
					
					errList.add("Error: Row "+s.get("resRowNumber")+" "+s.get("firstName")+" "+s.get("lastName")+" with Employee Number "+ s.get("employeeNumber") +
							": No Family Household Registration. ");
					
//					errList.add(res_s.get("firstName")+" "+res_s.get("lastName")+" with Employee Number "+ res_s.get("employeeNumber") +
//					": Record not exist in Registration File. ");
				}
			}
		}
		
		return errList;
	}

	private static HashMap<String, Object> convertListHashToHashMap(List<HashMap<String, Object>> arrListFromReserveFile) {
		
		HashMap<String, Object> mapList = new HashMap<String, Object>();
		
		for(HashMap<String, Object> res : arrListFromReserveFile) {
//			System.out.println(res);
			if(res.get("ModernactrlNumber").toString().trim() != "--") {
				String[] moCN = res.get("ModernactrlNumber").toString().trim().split(",");
				
				for(String ctrln : moCN) {
//					System.out.println(ctrln);
					
					String[] cn = ctrln.toString().trim().split("_");
					mapList.put(ctrln.toString().replaceAll(" ",""), cn[1]);
				}
			}
			
			if(res.get("CovovaxctrlNumber").toString().trim() != "--") {
				String[] cCN = res.get("CovovaxctrlNumber").toString().trim().split(",");
				
				for(String ctrln : cCN) {
//					System.out.println(ctrln);
					
					String[] cn = ctrln.toString().trim().split("_");
					mapList.put(ctrln.toString().replaceAll(" ",""), cn[1]);
				}
			}
			
		}
		
		return mapList;
	}

	private static ArrayList<String> excessCtrlNUmber(List<String> arrListFromRegFile, 
			List<HashMap<String, Object>> arrListFromReserveFile) {
		
		ArrayList<String> errList = new ArrayList<String>();
		HashMap<String, Object> mapList = new HashMap<String, Object>();
		HashMap<String, List<String>> mapArrListReg =  new HashMap<>();
		
		mapArrListReg = groupByEmployeeNumber(convertListToHashMap(arrListFromRegFile));
		System.out.println(mapArrListReg);
		
		int covCounter = 0;
		int modCounter = 0;
		
		for(String s : mapArrListReg.keySet()) {
			
			for(String r : mapArrListReg.get(s)) {
				String[] ctrln = r.trim().split("_");
				
				if(ctrln[2].contains("M")) {
					modCounter++;
				}else if(ctrln[2].contains("C")) {
					covCounter++;
				}
			}
//			System.out.println(s +" : "+ modCounter +" - "+ covCounter);
			
			mapList.put("employeeNumber", s);
			mapList.put("modernaqnty", modCounter);
			mapList.put("covovaxqnty", covCounter);
			
			for(HashMap<String, Object> res_s : arrListFromReserveFile) {
				
				if(res_s.get("employeeNumber").equals(mapList.get("employeeNumber"))) {
					if(Integer.parseInt(mapList.get("modernaqnty").toString()) != Integer.parseInt(res_s.get("modernaOrders").toString())) {
//						System.out.println(res_s.get("firstName")+" "+res_s.get("lastName")+" with Employee Number "+ res_s.get("employeeNumber") +
//								": Excess of Moderna Control Number. "+
//								mapList.get("modernaqnty").toString()+" in Registration - "+
//								res_s.get("modernaOrders").toString()+" in Reservation. ");
						
						errList.add(res_s.get("firstName")+" "+res_s.get("lastName")+" with Employee Number "+ res_s.get("employeeNumber") +
								": Excess of Moderna Control Number. "+
								mapList.get("modernaqnty").toString()+" in Registration - "+
								res_s.get("modernaOrders").toString()+" in Reservation. ");    
					}
					
					if(Integer.parseInt(mapList.get("covovaxqnty").toString()) != Integer.parseInt(res_s.get("covovaxOrders").toString())) {
//						System.out.println(res_s.get("firstName")+" "+res_s.get("lastName")+" with Employee Number "+ res_s.get("employeeNumber") +
//								": Excess of Covovax Control Number. "+
//								mapList.get("covovaxqnty").toString()+" in Registration - "+
//								res_s.get("covovaxOrders").toString()+" in Reservation. ");
						
						errList.add(res_s.get("firstName")+" "+res_s.get("lastName")+" with Employee Number "+ res_s.get("employeeNumber") +
								": Excess of Covovax Control Number. "+
								mapList.get("covovaxqnty").toString()+" in Registration - "+
								res_s.get("covovaxOrders").toString()+" in Reservation. ");
					}
				}
			}
			
//			System.out.println(mapList);
			modCounter=0;
			covCounter=0;
		}
		
//		System.out.println(arrListFromReserveFile);
		
		return errList;
	}

	private static HashMap<String, List<String>> groupByEmployeeNumber(HashMap<String, Object> listFromRegFile) {
		HashMap<String, List<String>> valuesMap = new HashMap<>();
//		 System.out.println("tt -- "+listFromRegFile);
		    for (String key : listFromRegFile.keySet()) {
		        Object val = listFromRegFile.get(key);
		        if (valuesMap.get(val) == null) {
		            List<String> values = new ArrayList<>();
		            values.add(key);
		            
		            valuesMap.put((String) val, values);
		        } else {
		            valuesMap.get(val).add(key);
		        }
		    }
//		System.out.println(valuesMap);
		    return valuesMap;

	}

	private static List<String> convertToList(List<HashMap<String, Object>> arrMap) {
		List<String> list = new ArrayList<String>();
		
		for(HashMap<String, Object> arr : arrMap) {
			if(arr.containsKey("isRed")) {
				if(Boolean.parseBoolean(arr.get("isRed").toString()) == false) {
					list.add(arr.get("controlNumber").toString());
				}
			}
		}
		
		for(String s: list) {
			
			
			if(s == "PAL_434411_M1") {
				System.out.println("convertToList --- "+s);
			}
		}
		
//		for(HashMap<String, Object> arr : arrMap) {
//			if(arr.containsKey("isRed")) {
//				if(Boolean.parseBoolean(arr.get("isRed").toString()) == false) {
//					list.add(arr.get("controlNumber").toString());
//				}
//			}else {
//				list.add(arr.get("controlNumber").toString());
//			}
//		}
		
		return list;
	}

	private static List<HashMap<String,Object>> getDataFromReservationFile(String excelFile) {
		List<String> ctrlList = new ArrayList<String>();
		
		List<HashMap<String, Object>> listsMap = new ArrayList<HashMap<String, Object>>();
		File file = new File(excelFile);
		FileInputStream fis;
		
		try {
			fis = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet spreadsheet = workbook.getSheetAt(0);
			
			Iterator < Row >  rowIterator = spreadsheet.iterator();
			
			while (rowIterator.hasNext()) {
		    	Row row = rowIterator.next();
		    	
		    	if (isBlankRow(row)) {
					continue;
				}
		    	
				if(row.getRowNum()==0){
					continue; //just skip the rows if row number is 0, 1, or 2
				}
		    	
		    	Iterator<Cell> cellIterator = row.cellIterator();
		    	HashMap<String, Object> mapList = new HashMap<String, Object>();
		    	
		    	while (cellIterator.hasNext()) {
		    		Cell cell = cellIterator.next();
		    		mapList.put("resRowNumber", row.getRowNum()+1);
		    		
		    		if(cell.getColumnIndex()==12) { //First Name
		    			switch (cell.getCellType()) {
			               case STRING:
			            	   mapList.put("firstName", cell.getStringCellValue());
			                  break;
			               case BLANK:
			            	   mapList.put("firstName", "--");
					              break;
						default:
							break;
			            }
		    		}
		    		
		    		if(cell.getColumnIndex()==11) { //Last Name
		    			switch (cell.getCellType()) {
			               case STRING:
			            	   mapList.put("lastName", cell.getStringCellValue());
			                  break;
			               case BLANK:
			            	   mapList.put("lastName", "--");
					              break;
						default:
							break;
			            }
		    		}
		    		
		    		if(cell.getColumnIndex()==10) { //employeeNumber
		    			switch (cell.getCellType()) {
			               case STRING:
			            	   mapList.put("employeeNumber", cell.getStringCellValue());
			                  break;
			               case NUMERIC:
			            	   mapList.put("employeeNumber", cell.getNumericCellValue());
				              break;
			               case BLANK:
			            	   mapList.put("employeeNumber", "--");
					              break;
						default:
							break;
			            }
		    		}
		    		
		    		if(cell.getColumnIndex()==18) { //For how many people are you reserving Moderna vaccines?
		    			switch (cell.getCellType()) {
			               case NUMERIC:
			                  mapList.put("modernaOrders", converterStringNum(cell.getNumericCellValue()));
			                  break;
			               case STRING: 
				              mapList.put("modernaOrders", converterStringNum(cell.getStringCellValue()));
				              break;
			               case BLANK:
			            	   mapList.put("modernaOrders", converterStringNum(0));
					              break;
			               case _NONE:
			            	   mapList.put("modernaOrders", converterStringNum(0));
			            	   break;
						default:
							break;
			            }
		    		}
		    		
		    		if(cell.getColumnIndex()==21) { //For how many people are you reserving Covovax (Novavax) vaccines?
		    			switch (cell.getCellType()) {
			               case NUMERIC:
			            	   mapList.put("covovaxOrders", converterStringNum(cell.getNumericCellValue()));
			                  break;
			               case STRING:
			            	   mapList.put("covovaxOrders", converterStringNum(cell.getStringCellValue()));
			            	   break;
			               case BLANK:
			            	   mapList.put("covovaxOrders", converterStringNum(0));
					           break;
						default:
							break;
			            }
		    		}
		    		
		    		String MCtrlnumber = null;
		    		String CCtrlnumber = null;
		    		
		    		if(cell.getColumnIndex()==25) { //moderna ctrl number
		    			switch (cell.getCellType()) {
			    			case NUMERIC:
				            	   mapList.put("ModernactrlNumber", "--");
				                  break;
			               case STRING: 
				              
				              String[] mCtrlnumber = cell.getStringCellValue().toString().trim().replaceAll(" ", "").split(",");
				              
				              if(cell.getStringCellValue().toString().trim().toLowerCase() == "na" || cell.getStringCellValue().toString().trim().toLowerCase() == "n/a" ) {
				            	  mapList.put("ModernactrlNumber", "--");
				              }else {
				            	  for(String arr : mCtrlnumber) {
					            	  String[] ctrlnumber = arr.toString().trim().split("_");
						              if(ctrlnumber.length >= 3) {
						            	  mapList.put("ModernactrlNumber", cell.getStringCellValue());
						            	  ctrlList.add(arr);
//						            	  System.out.println(arr);
						            	  
						              }
					              }
				              }
				              break;
			               case _NONE:
			            	  mapList.put("ModernactrlNumber", "--");
						      break;
			               case BLANK:
			            	  mapList.put("ModernactrlNumber", "--");
					          break;
						default:
							break;
			            }
		    		}
		    		
		    		if(cell.getColumnIndex()==26) { //covovax ctrl number
		    			switch (cell.getCellType()) {
		    			case NUMERIC:
			            	   mapList.put("CovovaxctrlNumber", "--");
			                  break;
			               case STRING:
					              String[] mCtrlnumber = cell.getStringCellValue().toString().trim().replaceAll(" ", "").split(",");
					          
					              if(cell.getStringCellValue().toString().trim().toLowerCase() == "na" || cell.getStringCellValue().toString() == "N/A" ) {
					            	  mapList.put("CovovaxctrlNumber", "--");
					              }else {
					            	  for(String arr : mCtrlnumber) {
						            	  String[] ctrlnumber = arr.toString().trim().split("_");
							              if(ctrlnumber.length >= 3) {
							            	  mapList.put("CovovaxctrlNumber", cell.getStringCellValue());
							            	  ctrlList.add(arr);
//							            	  System.out.println(arr);
//							            	  mapList.merge("ctrlNumber", cell.getStringCellValue().toString(), (oldValue, newValue) -> oldValue.toString() +","+ newValue.toString());
							              }
						              }
					              }
			            	   break;
			               case _NONE:
				               mapList.put("CovovaxctrlNumber", "--");
							   break;
			               case BLANK:
			            	   mapList.put("CovovaxctrlNumber", "--");
					           break;
						default:
							break;
			            }
		    		}
		    		mapList.put("ctrlNumber", "NAN");
		    		
		    	}
		    	listsMap.add(mapList);
			}
			workbook.close();
			
	    	for(HashMap<String, Object> s : listsMap) {
//	    		System.out.println(s);
	    		if(!s.containsKey("CovovaxctrlNumber")) {
	    			s.put("CovovaxctrlNumber", "--");
	    		}
	    		
	    		if(!s.containsKey("ModernactrlNumber")) {
	    			s.put("ModernactrlNumber", "--");
	    		}
	    	}
	    	
	    	for(HashMap<String, Object> s : listsMap) {
//	    		System.out.println(s);
	    	}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		return listsMap;
		
	}

	private static List<HashMap<String,Object>> getDataFromRegisterFile(String excelFile) {
		List<String> ctrlList = new ArrayList<String>();
		
		List<HashMap<String, Object>> listsMap = new ArrayList<HashMap<String, Object>>();
		File file = new File(excelFile);
		FileInputStream fis;
		
		try {
			fis = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet spreadsheet = workbook.getSheetAt(0);
			
			Iterator < Row >  rowIterator = spreadsheet.iterator();
	        
			while (rowIterator.hasNext()) {
		    	Row row = rowIterator.next();
		    	
				if (isBlankRow(row)) {
					continue;
				}
		    	
				if(row.getRowNum()==0 || row.getRowNum()==1){
					continue; //just skip the rows if row number is 0, 1, or 2
				}
		    	
		    	Iterator<Cell> cellIterator = row.cellIterator();
		    	List<Object> list = new ArrayList<Object>(); 
		    	HashMap<String, Object> mapList = new HashMap<String, Object>();
		    	
		    	while (cellIterator.hasNext()) {
		    		Cell cell = cellIterator.next();
		    		
		    		mapList.put("regRowNumber", row.getRowNum()+1);
		    		
		    		if(cell.getColumnIndex()==46) { //Control Number
		    			switch (cell.getCellType()) {
			               case STRING:
			            	   mapList.put("controlNumber", cell.getStringCellValue());
			            	   
//			            	   if(getFontColor(workbook, cell).equals("000000")) {
//			            		   mapList.put("isRed", false);
//			            	   }else {
//			            		   mapList.put("isRed", true);
//			            	   }
//			            	   
//			            	   mapList.put("isRed", false);
				              break;
			               case BLANK:
			            	   mapList.put("controlNumber", "--");
					              break;
						default:
							break;
			            }
		    		}
		    		
		    		if(cell.getColumnIndex()==5) { //Last Name
		    			switch (cell.getCellType()) {
			               case STRING:
				            	  mapList.put("lastName", cell.getStringCellValue());
				              break;
			               case BLANK:
			            	   mapList.put("controlNumber", "--");
					              break;
						default:
							break;
			            }
		    		}
		    		
		    		if(cell.getColumnIndex()==6) { //First Name
		    			switch (cell.getCellType()) {
			               case STRING:
				            	  mapList.put("firstName", cell.getStringCellValue());
				              break;
			               case BLANK:
			            	   mapList.put("controlNumber", "--");
					              break;
						default:
							break;
			            }
		    		}
		    		
		    		if(cell.getColumnIndex()==53) { //Company Name
		    			switch (cell.getCellType()) {
			               case STRING:
			            	   
				            	  mapList.put("companyCode", companyNameLookup(cell.getStringCellValue().toString()));
				              break;
			               case BLANK:
			            	   mapList.put("controlNumber", "--");
					              break;
						default:
							break;
			            }
		    		}
		    		
		    	}
		    	listsMap.add(mapList);
			}
			workbook.close();
			
//	    	for(HashMap<String, Object> s : listsMap) {
//	    		System.out.println(s);
//	    	}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return listsMap;
		
	}

	private static String companyNameLookup(String stringCellValue) {
		HashMap<String, String> cc = new HashMap<String, String>();
		
		cc.put("ALL", "All Seasons Realty Corp.");
		cc.put("APL", "Allianz-PNB Life Insurance, Inc. (APLII)");
//		cc.put("APL", "Allianz-PNB Life Insurance, Inc.");
		cc.put("ABI", "Asia Brewery, Inc. (ABI), Subsidiaries");
//		cc.put("ABI", "Asia Brewery, Inc. (ABI) and Subsidiaries");
//		cc.put("ABI", "ABI, its Subsidiaries, and Affiliates"); // 
		cc.put("BCH", "Basic Holdings Corp.");
		cc.put("CPH", "Century Park Hotel");
		cc.put("EPP", "Eton Properties Philippines, Inc. (EPPI), Subsidiaries");
//		cc.put("EPP", "Eton Properties Philippines, Inc. (Eton) and Subsidiaries");
//		cc.put("EPP", "EPPI and its Subsidiaries");
		cc.put("FFI", "Foremost Farms, Inc.");
		cc.put("FTC", "Fortune Tobacco Corp.");
		cc.put("GDC", "Grandspan Development Corp.");
		cc.put("HII", "Himmel Industries, Inc.");
		cc.put("LRC", "Landcom Realty Corp.");
		cc.put("LTG", "LT Group, Inc. (Parent Company)");
		cc.put("LTGC", "LTGC Directors");
//		cc.put("LTGC", "LT Group of Companies Directors");
//		cc.put("MAC", "MacroAsia Corp., Subsidiaries & Affiliates");
//		cc.put("MAC", "MAC, its Subsidiaries, and Affiliates");
		cc.put("MAC", "MacroAsia Corp., Subsidiaries and Affiliates");
//		cc.put("MAC", "MacroAsia Corp., Subsidiaries & Affiliates");
		cc.put("PAL", "Philippine Airlines, Inc. (PAL), Subsidiaries and Affiliates");
//		cc.put("PAL", "PAL, its Subsidiaries, and Affiliates");
		cc.put("PNB", "Philippine National Bank (PNB) and Subsidiaries");
//		cc.put("PNB", "PNB and its Subsidiaries");
		cc.put("PMI", "PMFTC Inc.");
		cc.put("RAP", "Rapid Movers & Forwarders, Inc.");
		cc.put("TYK", "Tan Yan Kee Foundation, Inc. (TYKFI)");
		cc.put("TDI", "Tanduay Distillers, Inc. (TDI) and subsidiaries");
//		cc.put("TDI", "TDI, its Subsidiaries, and Affiliates");
		cc.put("CHI", "Charter House Inc.");
//		cc.put("SPV", "Grandholdings Investments (SPV-AMC), Inc.");
//		cc.put("SPV", "Opal Portfolio Investments (SPV-AMC), Inc.");
		cc.put("SPV", "SPV-AMC Group");
//		cc.put("SPV", "SPV-AMC Group");
//		cc.put("SPV", "SPV Group");
		cc.put("TMC", "Topkick Movers Corporation");
		cc.put("UNI", "University of the East (UE)");
		cc.put("UER", "University of the East Ramon Magsaysay Memorial Medical Center (UERMMMC)");
//		cc.put("UER", "UERMMMC");
		cc.put("VMC", "Victorias Milling Company, Inc. (VMC)");
		cc.put("ZHI", "Zebra Holdings, Inc.");
		cc.put("STN", "Sabre Travel Network Phils., Inc.");
		cc.put("TMC", "Topkick Corp.");
//		cc.put("TMC", "Topkick Movers Corporation");
		
		return getKey(cc, stringCellValue);
	}
	
	public static <K, V> K getKey(Map<K, V> map, V value) {
        for (Map.Entry<K, V> entry : map.entrySet()) {
            if (value.equals(entry.getValue())) {
                return entry.getKey();
            }
        }
        return null;
    }

	private static boolean checkFile(String fileformat, String file) {
		File f = new File(file);
		if(f.isFile() && !f.isDirectory()) { 
			
			String filename = f.getName().toLowerCase();
			
			if(!filename.endsWith(fileformat)) {
				System.out.println(file + " is not valid excel format.");
				
				return false;
			}
		}else {
			System.out.println(file + " does not exist.");
			
			return false;
		}
		
		return true;
	}

	private static String getFileNameResult(String inReserveFile) {
		Date dNow = new Date();
		SimpleDateFormat ft = new SimpleDateFormat ("yyyy-MM-dd_(hh_mm_ss)");
		
		Path pathsrc = Paths.get(inReserveFile);
		Path srcFileName = pathsrc.getFileName();
		
		String[] scn;
		String companyNamefile = null;
		if(srcFileName.toString().contains("Daily")) {
			scn = srcFileName.toString().split("Daily");
			
			companyNamefile = scn[0].trim();
		}else if(srcFileName.toString().contains("Family")) {
			scn = srcFileName.toString().split("Family");
			
			companyNamefile = scn[0].replace("_","").trim();
		}
		
		return companyNamefile+"_crossReference_Err_Log_"+ft.format(dNow)+".txt";
	}
	
	private static boolean isBlankRow(Row row) {
        Cell cell;
        boolean result = true;
       
        for(int col = 0; col <= 72; col ++) {               
            cell = row.getCell(col);       
            /*if(row.getRowNum()>=8400) {
                System.out.println(cell + " - " + isCellEmpty(cell, false) );
            }*/
            if(!isCellEmpty(cell, false)) {
                result = false;
                break;                       
            }
        }
        return result;
    }
	
	private static Object converterStringNum(Object numericCellValue) {
		double d = Double.valueOf(numericCellValue.toString()).doubleValue();
		int orders = (int)d;
		
		
		return String.valueOf(orders);
	}
	
	private static boolean isCellEmpty(Cell cell, boolean checkForZero) {       
        if (cell == null) {
            return true;
        }   
        if (cell.getCellType() == CellType.BLANK) {
            return true;
        }   
        if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty()) {
            return true;
        }   
        if (checkForZero && cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == 0) {
            return true;
        }
        if (cell.getCellType() == CellType.FORMULA) {
            CellType cellType = cell.getCachedFormulaResultType();
            if(cellType == CellType.STRING && cell.getStringCellValue().trim().isEmpty()) {                                                       
                return true;                                                   
            }
           
            if(checkForZero && cellType == CellType.NUMERIC && cell.getNumericCellValue() == 0) {                                                       
                return true;                                                   
            }
        }
        return false;
   }
	
	public static String getFontColor(Workbook workbook, Cell cell) {

	    final CellStyle cellStyle = cell.getCellStyle();
	    final short fontIndex = (short) cellStyle.getFontIndex();
	    final XSSFFont font = (XSSFFont) workbook.getFontAt(fontIndex);
	
	    // Font color
	    final short colorIndex = font.getColor();
	    XSSFFont xssfFont = font;
	    String myFontColor = null;
	    
        XSSFColor color = xssfFont.getXSSFColor();
        if (color != null) {
            String argbHex = color.getARGBHex();
            if (argbHex != null) {
                myFontColor = argbHex.substring(2);
            }
        }
        
        return myFontColor;
	}
	
	private static HashMap<String, Object> convertListToHashMap(List<String> arrListFromRegFile) {
		HashMap<String, Object> mapList = new HashMap<String, Object>();
		System.out.println("----"+arrListFromRegFile);
		for(String s : arrListFromRegFile) {
			String[] regCtrlNumber = s.toString().trim().split("_");
//			System.out.println(regCtrlNumber.length);
//			if(regCtrlNumber.length == )
			
			mapList.put(s, regCtrlNumber[1]);
		}
		
		return mapList;
	}
	
	
	
	public static String getInFilePath() {
		return inFilePath;
	}

	public static void setInFilePath(String inFilePath) {
		Application.inFilePath = inFilePath;
	}

	public static String getInRegisterFile() {
		return inRegisterFile;
	}

	public static void setInRegisterFile(String inRegisterFile) {
		Application.inRegisterFile = inRegisterFile;
	}

	public static String getInReserveFile() {
		return inReserveFile;
	}

	public static void setInReserveFile(String inReserveFile) {
		Application.inReserveFile = inReserveFile;
	}

	public static String getOutResultFile() {
		return outResultFile;
	}

	public static void setOutResultFile(String outResultFile) {
		Application.outResultFile = outResultFile;
	}

	
}
