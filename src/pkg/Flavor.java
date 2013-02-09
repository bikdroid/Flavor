package pkg;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;

import jxl.write.WriteException;


public class Flavor {

	public static void main(String[] args) throws Exception {

		String sourceDir = "D:\\Lectures";
		String targetFile = "D:\\Lectures\\Scanned.xls";
		Boolean processChilds = true;
		
		ArrayList<String> textFileList = getTextFileList(sourceDir, processChilds);
		HashMap<String, String> textFileMap = getTextFileMap(textFileList);
		sendToExcel(textFileMap, targetFile);
	}

	private static void sendToExcel(HashMap<String, String> textFileMap, String targetFile) throws WriteException, IOException {
		
		System.out.println("Writing to excel file: " + targetFile);
		WriteExcel toExcel = new WriteExcel(textFileMap, targetFile);
		toExcel.write();
		System.out.println("Write Complete.");
	}

	private static ArrayList<String> getTextFileList(String dirName, Boolean processChilds) throws IOException {
		File dir = new File(dirName);
		ExtFilter txtOnly = new ExtFilter();
		File[] fileList = dir.listFiles(txtOnly);
		ArrayList<String> textFileList = new ArrayList<String>();
		for (File file : fileList) {
			if (file.isFile())
				textFileList.add(file.getAbsolutePath());
		}

		/* If you want to scan child folders as well... */
		if(processChilds==true) {
			File[] allFileList = dir.listFiles();
			for (File file : allFileList) {
				if (file.isDirectory()) {
					textFileList.addAll(getTextFileList(file.getAbsolutePath(), processChilds));
				}
			}
		}
		return textFileList;
	}

	private static HashMap<String, String> getTextFileMap(ArrayList<String> fileList)
			throws Exception {

		StringBuilder sbr;
		String line, text;
		HashMap<String, String> textMap = new HashMap<String, String>();

		if (fileList.isEmpty()) {
			System.out.println("We could not find any text file.");
			return null;
		}

		for (String file : fileList) {

			System.out.println("\nProcessing: " + file);
	
			BufferedReader br = new BufferedReader(new FileReader(file));
			sbr = new StringBuilder();
			while ((line = br.readLine()) != null) {
				sbr.append(line + " ");
			}

			text = sbr.toString().replaceAll("\\s+", " ");
			System.out.println(file + " --> " + text);
			textMap.put(file, text);
			System.out.println("Processed: " + file);
			br.close();
		}
		return textMap;
	}

}
