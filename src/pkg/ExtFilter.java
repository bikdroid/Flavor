package pkg;
import java.io.*;

public class ExtFilter implements FilenameFilter {
	@Override
	public boolean accept(File currentDirectory, String fileName) {
		/* Lower case the filename for easy checking */
		if(fileName.toLowerCase().endsWith(".txt"))
			return true;
		else 
			return false;
	}
}
