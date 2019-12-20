package xLsToXmL;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.JFileChooser;
import javax.swing.JTextArea;
import javax.swing.filechooser.FileNameExtensionFilter;


public class ConvertListener implements ActionListener {

    //private JTextArea origin;
    private JTextArea destination;

    public ConvertListener(JTextArea destination) {
        //this.origin = origin;
        this.destination = destination;
    }


    @Override
    public void actionPerformed(ActionEvent ae) {
    	String filePath = "";
    	String directoryPath = "";
    	String fileName = "";
    	JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "xls & xlsx", "xls", "xlsx");
        chooser.setFileFilter(filter);
        int returnVal = chooser.showOpenDialog(null);
        if(returnVal == JFileChooser.APPROVE_OPTION) {
            System.out.println("You chose to open this file: " +
            		chooser.getSelectedFile().getName().replaceFirst("[.][^.]+$", ""));
//declares the filePath of the original file to be used in the filereader, then the directory path to where the xml file is going to be created
            filePath = chooser.getSelectedFile().getAbsolutePath();
            directoryPath = chooser.getSelectedFile().getParent();
            //there's supposed to be an apache commons fileNameUtils, or using java.nio.file.Files/Path way of doing this, but for now I use regex to remove the extension from the file name
            fileName = chooser.getSelectedFile().getName().replaceFirst("[.][^.]+$", "");
            this.destination.setText("Converted xml file saved to: " + directoryPath + "\\" + fileName + ".xml");
        }
    	FileReader.converterMethod(filePath, directoryPath, fileName);
    }
}